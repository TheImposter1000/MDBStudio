using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace MDBStudio;

public partial class ExcelImportDialog : Window
{
    private readonly string _fileName;
    private readonly IReadOnlyList<string> _tables;
    private readonly bool _isCsv;
    private int _generation;

    public DataTable? SelectedSheet { get; private set; }
    public bool CreateNewTable => TargetCombo.SelectedIndex == 0;
    public string Destination => CreateNewTable
        ? NewTableBox.Text.Trim()
        : Convert.ToString(TargetCombo.SelectedItem) ?? "";

    public ExcelImportDialog(string filename, IReadOnlyList<string> tables)
    {
        InitializeComponent();
        _fileName = filename;
        _tables = tables;
        _isCsv = string.Equals(Path.GetExtension(filename), ".csv", StringComparison.OrdinalIgnoreCase);
        SourceLabel.Text = filename;
        NewTableBox.Text = DefaultName(Path.GetFileNameWithoutExtension(filename));
        TargetCombo.ItemsSource = new[] { "(Create a new MDB table)" }.Concat(tables).ToArray();
        TargetCombo.SelectedIndex = 0;
        SheetCombo.ItemsSource = _isCsv
            ? new[] { CsvTransfer.FileLabel }
            : ExcelTransfer.Sheets(filename);
        if (SheetCombo.Items.Count > 0) SheetCombo.SelectedIndex = 0;
    }

    private static string DefaultName(string file)
    {
        var cleaned = new string(file.Select(ch =>
            ch is ']' or '[' or '\0' ? '_' : ch).ToArray()).Trim();
        if (cleaned.StartsWith("MSys", StringComparison.OrdinalIgnoreCase))
            cleaned = "Imported_" + cleaned;
        return string.IsNullOrWhiteSpace(cleaned) ? "ExcelImport" :
            cleaned[..Math.Min(64, cleaned.Length)];
    }

    private async void SheetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SheetCombo.SelectedItem is not string name) return;
        int version = ++_generation;
        SelectedSheet = null;
        ImportButton.IsEnabled = false;
        PreviewInfo.Text = _isCsv ? "Reading CSV file..." : "Reading Excel worksheet...";
        try
        {
            var data = await Task.Run(() => _isCsv
                ? CsvTransfer.ReadFile(_fileName)
                : ExcelTransfer.ReadSheet(_fileName, name));
            if (version != _generation) return;
            SelectedSheet = data;
            var preview = data.Clone();
            foreach (var row in data.Rows.Cast<DataRow>().Take(30))
                preview.ImportRow(row);
            PreviewGrid.ItemsSource = preview.DefaultView;
            RefreshPreviewInfo();
        }
        catch (Exception ex)
        {
            if (version != _generation) return;
            PreviewGrid.ItemsSource = null;
            PreviewInfo.Text = "Unable to read worksheet: " + ex.Message;
            MessageBox.Show(this, ex.Message, "Data import", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (version == _generation) ImportButton.IsEnabled = SelectedSheet is not null;
        }
    }

    private void TargetChanged(object sender, SelectionChangedEventArgs e)
    {
        if (NewTableBox is null) return;
        NewTableBox.IsEnabled = TargetCombo.SelectedIndex == 0;
        RefreshPreviewInfo();
    }

    private void RefreshPreviewInfo()
    {
        if (SelectedSheet is null || PreviewInfo is null) return;
        PreviewInfo.Text = $"{SelectedSheet.Rows.Count:N0} data rows and " +
            $"{SelectedSheet.Columns.Count:N0} columns. " +
            (CreateNewTable
                ? "Create a new MDB table. Columns are imported as text so identifiers with leading zeros are preserved; an AutoNumber primary key is added."
                : "Append records to the selected existing table. Columns are matched by their Excel header names; unmatched headings are ignored. Database field constraints still apply.");
    }

    private void Import_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedSheet is null) return;
        string destination = Destination;
        if (string.IsNullOrWhiteSpace(destination) || destination.Length > 64 ||
            destination.Contains(']') || destination.StartsWith("MSys", StringComparison.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "Enter a valid table name of up to 64 characters.",
                "Invalid MDB table name", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (CreateNewTable && _tables.Contains(destination, StringComparer.OrdinalIgnoreCase))
        {
            MessageBox.Show(this, "That table already exists. Choose a new name.",
                "Table already exists", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (MessageBox.Show(this,
            $"Import {SelectedSheet.Rows.Count:N0} rows from '{Path.GetFileName(_fileName)}' " +
            $"{(CreateNewTable ? "into a NEW table" : "into existing table")} '{destination}'?\n\n" +
            "MDB Studio will create a backup, then commit the import as one operation. " +
            "This action is separate from Save Changes. Continue?",
            "Confirm data import", MessageBoxButton.YesNo, MessageBoxImage.Question,
            MessageBoxResult.No) != MessageBoxResult.Yes)
            return;
        DialogResult = true;
    }
}
