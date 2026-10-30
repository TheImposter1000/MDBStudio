using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace MDBStudio;

public partial class MainWindow : Window
{
    private AccessSession? _session;
    private DataTable? _table;
    private string? _currentName;
    private bool _suppressSelection;

    public MainWindow() => InitializeComponent();

    private void Status(string message) => StatusLabel.Text = message;

    private static void Error(string heading, Exception ex) =>
        MessageBox.Show(ex.Message, heading, MessageBoxButton.OK, MessageBoxImage.Error);

    private bool DiscardPending()
    {
        RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);
        if (_table?.GetChanges() is null) return true;
        return MessageBox.Show("Discard unsaved changes?", "MDB Studio",
            MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private void Open_Click(object sender, RoutedEventArgs e)
    {
        if (!DiscardPending()) return;
        var dialog = new OpenFileDialog
        {
            Filter = "Access databases (*.mdb;*.accdb)|*.mdb;*.accdb",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            AccessSession next;
            try
            {
                next = new AccessSession(dialog.FileName);
            }
            catch (DatabasePasswordRequiredException)
            {
                var prompt = new DatabasePasswordDialog(Path.GetFileName(dialog.FileName))
                {
                    Owner = this
                };
                if (prompt.ShowDialog() != true)
                {
                    Status("Database opening cancelled.");
                    return;
                }
                // The password stays in this process and is never saved in settings or logs.
                next = new AccessSession(dialog.FileName, prompt.EnteredPassword);
            }
            var names = next.ListTables();
            _session = next;
            _table = null;
            _currentName = null;
            _suppressSelection = true;
            TableList.ItemsSource = names;
            TableList.SelectedIndex = -1;
            _suppressSelection = false;
            RecordsGrid.ItemsSource = null;
            QueryGrid.ItemsSource = null;
            FileLabel.Text = dialog.FileName;
            TableLabel.Text = "Choose a table";
            KeyLabel.Text = "";
            SaveButton.IsEnabled = false;
            BackupButton.IsEnabled = true;
            ExportButton.IsEnabled = false;
            Status("Opened " + Path.GetFileName(dialog.FileName) + ". " + names.Count + " tables found.");
        }
        catch (Exception ex)
        {
            _suppressSelection = false;
            Error("Could not open database", ex);
        }
    }

    private void TableList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressSelection || _session is null || TableList.SelectedItem is not string name) return;
        if (_currentName == name) return;

        if (!DiscardPending())
        {
            _suppressSelection = true;
            TableList.SelectedItem = _currentName;
            _suppressSelection = false;
            return;
        }
        LoadTable(name);
    }

    private void LoadTable(string name)
    {
        try
        {
            var table = _session!.Load(name);
            _table = table;
            _currentName = name;
            RecordsGrid.ItemsSource = table.DefaultView;
            var writable = table.PrimaryKey.Length != 0;
            RecordsGrid.IsReadOnly = !writable;
            RecordsGrid.CanUserAddRows = writable;
            RecordsGrid.CanUserDeleteRows = writable;
            SaveButton.IsEnabled = writable;
            ExportButton.IsEnabled = true;
            TableLabel.Text = name;
            KeyLabel.Text = writable ? "Edit enabled (primary key detected)" : "Read only (no primary key)";
            Status("Loaded " + table.Rows.Count + " records (maximum " + AccessSession.PreviewRowLimit + ").");
        }
        catch (Exception ex) { Error("Could not load table", ex); }
    }

    private void Reload_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || !DiscardPending()) return;
        LoadTable(_currentName);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null || _table is null || _currentName is null) return;
        try
        {
            if (!RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true) ||
                !RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true))
                throw new InvalidOperationException("Complete or cancel the active cell edit before saving.");
            if (_table.GetChanges() is null)
            {
                Status("No pending changes.");
                return;
            }
            int changed = _session.Save(_table);
            string current = _currentName;
            LoadTable(current);  // Re-read autogenerated IDs and server-side defaults.
            Status("Saved " + changed + " changes. Backup exists in the MDBStudio Backups folder.");
        }
        catch (Exception ex) { Error("Save failed. Changes were rolled back when possible.", ex); }
    }

    private void Backup_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null) return;
        try { Status("Backup created: " + _session.Backup()); }
        catch (Exception ex) { Error("Backup failed", ex); }
    }

    private void Query_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null)
        {
            MessageBox.Show("Open a database first.", "MDB Studio");
            return;
        }
        try
        {
            var results = _session.Select(SqlInput.Text);
            QueryGrid.ItemsSource = results.DefaultView;
            Status("Query returned " + results.Rows.Count + " rows (maximum " + AccessSession.PreviewRowLimit + ").");
        }
        catch (Exception ex) { Error("Query failed", ex); }
    }

    private static string CsvCell(object value)
    {
        var raw = value == DBNull.Value ? "" :
            value is DateTime dt ? dt.ToString("O", CultureInfo.InvariantCulture) :
            Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
        // Avoid Excel formula execution when opening exported CSV.
        if (raw.Length > 0 && "=+-@\t\r".Contains(raw[0])) raw = "'" + raw;
        return "\"" + raw.Replace("\"", "\"\"") + "\"";
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || _currentName is null) return;
        var dialog = new SaveFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv",
            FileName = _currentName + ".csv",
            OverwritePrompt = true
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            using var writer = new StreamWriter(dialog.FileName, false, new UTF8Encoding(true));
            var columns = _table.Columns.Cast<DataColumn>().ToArray();
            writer.WriteLine(string.Join(",", columns.Select(c => CsvCell(c.ColumnName))));
            foreach (DataRow row in _table.Rows)
            {
                if (row.RowState == DataRowState.Deleted) continue;
                writer.WriteLine(string.Join(",", columns.Select(c => CsvCell(row[c]))));
            }
            Status("Exported loaded rows to " + dialog.FileName);
        }
        catch (Exception ex) { Error("Export failed", ex); }
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
        if (!DiscardPending()) e.Cancel = true;
    }
}
