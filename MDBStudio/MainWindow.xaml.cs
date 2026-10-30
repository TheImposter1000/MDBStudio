using Microsoft.Win32;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MDBStudio;

public partial class MainWindow : Window
{
    private AccessSession? _session;
    private DataTable? _table;
    private string? _currentName;
    private bool _suppressSelection;
    private bool _suppressPageSizeChange;
    private int _pageSize = AccessSession.DefaultPageSize;
    private int _pageIndex;
    private long _totalRecords;
    private IReadOnlyList<string> _allTables = Array.Empty<string>();
    private IReadOnlyList<DataRow>? _rightClickSelection;

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
            _pageIndex = 0;
            _totalRecords = 0;
            _suppressSelection = true;
            _allTables = names;
            TableSearchBox.Clear();
            TableList.ItemsSource = names;
            TableList.SelectedIndex = -1;
            _suppressSelection = false;
            RecordsGrid.ItemsSource = null;
            QueryGrid.ItemsSource = null;
            FileLabel.Text = dialog.FileName;
            DatabaseTitle.Text = Path.GetFileName(dialog.FileName);
            RecordSearchBox.Clear();
            TableLabel.Text = "Choose a table";
            KeyLabel.Text = "";
            SaveButton.IsEnabled = false;
            BackupButton.IsEnabled = true;
            ExportButton.IsEnabled = false;
            PreviousPageButton.IsEnabled = false;
            NextPageButton.IsEnabled = false;
            PageInfoLabel.Text = "Choose a table";
            UpdateActionStates();
            StatusRight.Text = "Database connected";
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
        LoadTable(name, 0);
    }

    private void LoadTable(string name, int requestedPage)
    {
        try
        {
            var page = _session!.LoadPage(name, requestedPage, _pageSize);
            var table = page.Rows;
            _table = table;
            _currentName = name;
            _pageIndex = page.PageIndex;
            _totalRecords = page.TotalRecords;
            RecordsGrid.ItemsSource = table.DefaultView;
            RecordSearchBox.Clear();
            table.RowChanged += (_, _) => Dispatcher.BeginInvoke(UpdateActionStates);
            table.RowDeleted += (_, _) => Dispatcher.BeginInvoke(UpdateActionStates);

            bool writable = table.PrimaryKey.Length != 0;
            RecordsGrid.IsReadOnly = !writable;
            RecordsGrid.CanUserAddRows = writable;
            RecordsGrid.CanUserDeleteRows = writable;
            SaveButton.IsEnabled = writable;
            ExportButton.IsEnabled = true;
            UpdateActionStates();
            TableLabel.Text = name;
            KeyLabel.Text = writable ? "Edit enabled (primary key detected)" : "Read only (no primary key)";
            UpdatePageControls(table.Rows.Count);
            Status($"Loaded {table.Rows.Count:N0} rows from {name}. Total: {_totalRecords:N0}. No 2,000-row ceiling.");
        }
        catch (Exception ex) { Error("Could not load table", ex); }
    }

    private void UpdatePageControls(int pageRows)
    {
        PreviousPageButton.IsEnabled = _currentName is not null && _pageSize > 0 && _pageIndex > 0;
        NextPageButton.IsEnabled = _currentName is not null && _pageSize > 0 &&
            ((long)_pageIndex + 1) * _pageSize < _totalRecords;

        if (_currentName is null)
        {
            PageInfoLabel.Text = "Choose a table";
            return;
        }

        if (_totalRecords == 0)
        {
            PageInfoLabel.Text = "0 records";
            return;
        }
        if (_pageSize == 0)
        {
            PageInfoLabel.Text = $"All {_totalRecords:N0} records";
            return;
        }
        long start = (long)_pageIndex * _pageSize + 1;
        long finish = start + pageRows - 1;
        long pages = (_totalRecords - 1) / _pageSize + 1;
        PageInfoLabel.Text = $"Records {start:N0} to {finish:N0} of {_totalRecords:N0} | Page {_pageIndex + 1:N0} of {pages:N0}";
    }

    private void PreviousPage_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || _pageIndex == 0 || !DiscardPending()) return;
        LoadTable(_currentName, _pageIndex - 1);
    }

    private void NextPage_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || _pageSize == 0 ||
            ((long)_pageIndex + 1) * _pageSize >= _totalRecords || !DiscardPending()) return;
        LoadTable(_currentName, _pageIndex + 1);
    }

    private void PageSize_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_suppressPageSizeChange || _session is null || _currentName is null ||
            PageSizeCombo.SelectedItem is not ComboBoxItem selection ||
            !int.TryParse(Convert.ToString(selection.Tag), out var newSize) || newSize == _pageSize)
            return;

        if (!DiscardPending())
        {
            _suppressPageSizeChange = true;
            try
            {
                foreach (var item in PageSizeCombo.Items.OfType<ComboBoxItem>())
                {
                    if (int.TryParse(Convert.ToString(item.Tag), out var tagSize) &&
                        tagSize == _pageSize)
                    {
                        PageSizeCombo.SelectedItem = item;
                        break;
                    }
                }
            }
            finally { _suppressPageSizeChange = false; }
            return;
        }

        _pageSize = newSize;
        LoadTable(_currentName, 0);
    }

    private void Reload_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || !DiscardPending()) return;
        LoadTable(_currentName, _pageIndex);
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
            LoadTable(current, _pageIndex);  // Re-read autogenerated IDs and server-side defaults.
            Status("Saved " + changed + " changes. Backup exists in the MDBStudio Backups folder.");
            UpdateActionStates();
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
            Status("Query returned " + results.Rows.Count.ToString("N0") + " rows.");
        }
        catch (Exception ex) { Error("Query failed", ex); }
    }

    private void Export_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null || _currentName is null) return;
        if (_table?.GetChanges() is not null &&
            MessageBox.Show("The export includes all SAVED database records, not unsaved edits. Continue?",
                "Export entire table", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            return;

        var dialog = new SaveFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv",
            FileName = _currentName + ".csv",
            OverwritePrompt = true
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            long count = _session.ExportTableCsv(_currentName, dialog.FileName);
            Status($"Exported all {count:N0} saved records from {_currentName} to {dialog.FileName}.");
        }
        catch (Exception ex) { Error("CSV export failed", ex); }
    }

    private async void CheckUpdates_Click(object sender, RoutedEventArgs e)
    {
        CheckUpdatesButton.IsEnabled = false;
        Status("Checking the official MDB Studio GitHub releases...");
        try
        {
            UpdateResult result = await UpdateChecker.CheckAsync();
            if (!result.UpdateAvailable)
            {
                Status($"MDB Studio {result.CurrentTag} is current; latest release is {result.LatestTag}.");
                MessageBox.Show($"You're using {result.CurrentTag}. The newest published release is {result.LatestTag}.",
                    "MDB Studio updates", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Status($"Update found: {result.LatestTag}. Open GitHub to review the release.");
            var answer = MessageBox.Show(
                $"An update is available.\n\nCurrent: {result.CurrentTag}\nLatest: {result.LatestTag}\n\n" +
                "The release may be an unverified preview. Always respect Windows Defender warnings. " +
                "MDB Studio will not download, install, or execute updates automatically.\n\n" +
                "Open the official GitHub release page in your browser?",
                "MDB Studio update available",
                MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (answer == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo(result.LatestUrl.AbsoluteUri)
                {
                    UseShellExecute = true
                });
            }
        }
        catch (Exception ex)
        {
            Status("Unable to check for updates. Check your internet connection.");
            Error("Update check failed", ex);
        }
        finally
        {
            CheckUpdatesButton.IsEnabled = true;
        }
    }


    private void UpdateActionStates()
    {
        bool open = _table is not null;
        bool writable = open && _table!.PrimaryKey.Length > 0;
        bool selected = RecordsGrid.SelectedItems.OfType<DataRowView>()
            .Any(row => row.Row.RowState != DataRowState.Deleted);
        CopyButton.IsEnabled = open && selected;
        PasteButton.IsEnabled = writable;
        DuplicateButton.IsEnabled = writable && selected;
        DeleteButton.IsEnabled = writable && selected;
        if (open)
        {
            int dirty = _table!.Rows.Cast<DataRow>()
                .Count(row => row.RowState != DataRowState.Unchanged);
            StatusRight.Text = dirty > 0
                ? $"{dirty:N0} unsaved changes | {_totalRecords:N0} saved rows"
                : $"All changes saved | {_totalRecords:N0} records";
        }
    }

    private IReadOnlyList<DataRow> SelectedRows(object? source = null)
    {
        // WPF may change the row selection when a context menu opens.
        // Preserve the complete pre-click selection when the menu was
        // opened on an already selected row.
        if (source is MenuItem && _rightClickSelection is { Count: > 0 } context)
        {
            _rightClickSelection = null;
            return context.Where(row => row.RowState != DataRowState.Deleted).ToArray();
        }
        _rightClickSelection = null;
        return RecordsGrid.SelectedItems.OfType<DataRowView>()
            .Select(view => view.Row)
            .Where(row => row.RowState != DataRowState.Deleted)
            .ToArray();
    }

    private void RecordsGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _rightClickSelection = null;
        if (e.OriginalSource is not DependencyObject element) return;
        DependencyObject? current = element;
        while (current is not null)
        {
            if (current is DataGridRow row)
            {
                if (row.IsSelected)
                {
                    _rightClickSelection = RecordsGrid.SelectedItems
                        .OfType<DataRowView>()
                        .Select(item => item.Row)
                        .Where(item => item.RowState != DataRowState.Deleted)
                        .ToArray();
                }
                break;
            }
            current = VisualTreeHelper.GetParent(current);
        }
    }

    private void RecordsSelection_Changed(object sender, SelectedCellsChangedEventArgs e) =>
        UpdateActionStates();

    private void RecordsGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.C)
        {
            CopyRecords_Click(sender, e);
            e.Handled = true;
        }
        else if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.V)
        {
            PasteRecords_Click(sender, e);
            e.Handled = true;
        }
        else if (e.Key == Key.Delete && (Keyboard.Modifiers & ModifierKeys.Control) == 0)
        {
            DeleteRecords_Click(sender, e);
            e.Handled = true;
        }
    }

    private void CopyRecords_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;
        try
        {
            RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var rows = SelectedRows(sender);
            if (rows.Count == 0) { Status("Select one or more complete rows to copy."); return; }
            if (rows.Count > 10000)
                throw new InvalidOperationException("Copy no more than 10,000 records at once.");
            int count = RecordClipboard.Copy(_table, rows);
            Status($"Copied {count:N0} complete records to the clipboard (Ctrl+C).");
        }
        catch (Exception ex) { Error("Copy records", ex); }
    }

    private void PasteRecords_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;
        try
        {
            RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var ids = _session!.ReadPasteIdentityCatalog(_table);
            var result = RecordClipboard.Paste(_table, ids, false);
            Status($"Pasted {result.Added:N0} new records. Generated {result.GeneratedIds:N0} new IDs and " +
                $"{result.GeneratedNames:N0} new Recordname values. Review rows, then Save Changes.");
            UpdateActionStates();
        }
        catch (Exception ex) { Error("Paste records", ex); }
    }

    private void DuplicateRecords_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null) return;
        try
        {
            RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var selected = SelectedRows(sender);
            var ids = _session!.ReadPasteIdentityCatalog(_table);
            var result = RecordClipboard.Paste(_table, ids, true, selected);
            Status($"Duplicated {result.Added:N0} records. Generated {result.GeneratedIds:N0} IDs and " +
                $"{result.GeneratedNames:N0} Recordname values. Save Changes to commit.");
            UpdateActionStates();
        }
        catch (Exception ex) { Error("Duplicate records", ex); }
    }

    private void DeleteRecords_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || _table.PrimaryKey.Length == 0) return;
        var rows = SelectedRows(sender);
        if (rows.Count == 0) return;
        if (MessageBox.Show(this,
            $"Mark {rows.Count:N0} selected records for deletion?\n\n" +
            "Nothing is removed from the MDB file until you click Save Changes.",
            "Delete records", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            return;
        foreach (var row in rows)
        {
            if (row.RowState == DataRowState.Added) _table.Rows.Remove(row);
            else row.Delete();
        }
        Status($"{rows.Count:N0} records marked for deletion. Save Changes to apply.");
        UpdateActionStates();
    }

    private void TableSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (TableList is null || _allTables.Count == 0) return;
        var query = TableSearchBox.Text.Trim();
        _suppressSelection = true;
        try
        {
            TableList.ItemsSource = _allTables
                .Where(name => name.Contains(query, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (_currentName is not null && TableList.Items.Contains(_currentName))
                TableList.SelectedItem = _currentName;
        }
        finally { _suppressSelection = false; }
    }

    private void RecordSearch_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_table is null || RecordSearchBox is null) return;
        string query = RecordSearchBox.Text.Trim();
        if (query.Length == 0)
        {
            _table.DefaultView.RowFilter = "";
            return;
        }
        string escaped = query.Replace("'", "''")
            .Replace("[", "[[]").Replace("]", "[]]")
            .Replace("%", "[%]").Replace("*", "[*]");
        var fields = _table.Columns.Cast<DataColumn>()
            .Where(c => c.DataType == typeof(string) && !c.ColumnName.Contains(']'))
            .Select(c => $"[{c.ColumnName}] LIKE '%{escaped}%'").ToArray();
        if (fields.Length == 0) return;
        try
        {
            _table.DefaultView.RowFilter = string.Join(" OR ", fields);
        }
        catch (SyntaxErrorException)
        {
            _table.DefaultView.RowFilter = "";
            Status("Search contains unsupported characters.");
        }
    }

    private void Window_Closing(object sender, CancelEventArgs e)
    {
        if (!DiscardPending()) e.Cancel = true;
    }
}
