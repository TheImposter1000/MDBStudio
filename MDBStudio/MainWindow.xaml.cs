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
    private IReadOnlyList<DataGridCellInfo>? _rightClickCells;
    private bool _applyingCellChanges;

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
            table.RowChanged += (_, _) => { if (!_applyingCellChanges) Dispatcher.BeginInvoke(UpdateActionStates); };
            table.RowDeleted += (_, _) => { if (!_applyingCellChanges) Dispatcher.BeginInvoke(UpdateActionStates); };

            bool writable = table.PrimaryKey.Length != 0;
            RecordsGrid.IsReadOnly = !writable;
            RecordsGrid.CanUserAddRows = writable;
            RecordsGrid.CanUserDeleteRows = false; // Explicit row deletion only; Delete key clears selected cells.
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
        Status("Checking for MDB Studio updates on the official GitHub repository...");
        bool installerStarted = false;
        try
        {
            var result = await UpdateChecker.CheckAsync();
            if (!result.UpdateAvailable)
            {
                Status($"MDB Studio {result.CurrentTag} is up to date.");
                MessageBox.Show(this, $"You're using {result.CurrentTag}. This is the latest release.",
                    "MDB Studio updates", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (result.ZipUrl is null || result.ChecksumUrl is null)
            {
                throw new InvalidOperationException(
                    "This release is missing an official verified ZIP or checksum. " +
                    "The application will not install an unverified package.");
            }

            var proceed = MessageBox.Show(this,
                $"MDB Studio {result.LatestTag} is available.\n\nInstalled: {result.CurrentTag}\n" +
                $"Available: {result.LatestTag}\n\n" +
                "Install this update directly into your current MDB Studio folder? " +
                "The application will close, update its own files, and reopen automatically. " +
                "Your MDB database files are not part of the update.\n\n" +
                "Do not bypass Windows Defender warnings.",
                "Install MDB Studio update", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (proceed != MessageBoxResult.Yes)
            {
                Status("Update canceled.");
                return;
            }

            if (!RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true) ||
                !RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true))
                throw new InvalidOperationException("Finish editing the active record before updating.");
            if (_table?.GetChanges() is not null)
            {
                MessageBox.Show(this,
                    "You have unsaved database changes. Save your records first, then check for updates again. " +
                    "MDB Studio will not discard pending edits during an update.",
                    "Save changes before updating", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            Status($"Downloading and verifying {result.LatestTag}. No files are being replaced yet...");
            var staged = await UpdateChecker.DownloadAndVerifyAsync(result);
            var launch = new ProcessStartInfo(staged.UpdaterPath)
            {
                WorkingDirectory = staged.WorkDirectory,
                UseShellExecute = false
            };
            launch.ArgumentList.Add("--apply");
            launch.ArgumentList.Add(Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture));
            launch.ArgumentList.Add(Path.TrimEndingDirectorySeparator(AppContext.BaseDirectory));
            launch.ArgumentList.Add(staged.ArchivePath);
            launch.ArgumentList.Add(staged.Sha256);

            if (Process.Start(launch) is null)
                throw new InvalidOperationException("Unable to start the update installer.");
            installerStarted = true;
            Status("Verified update downloaded. Closing to install, then relaunching MDB Studio...");
            Close();
        }
        catch (Exception ex)
        {
            Status("Update could not be installed; existing files were not intentionally changed.");
            Error("MDB Studio update failed", ex);
        }
        finally
        {
            if (!installerStarted) CheckUpdatesButton.IsEnabled = true;
        }
    }

    private void UpdateActionStates()
    {
        bool open = _table is not null;
        bool writable = open && _table!.PrimaryKey.Length > 0;
        int cellCount = RecordsGrid.SelectedCells.Count;
        CopyButton.IsEnabled = open && cellCount > 0;
        PasteButton.IsEnabled = writable && cellCount > 0;
        FillDownButton.IsEnabled = writable && cellCount > 1;
        DeleteButton.IsEnabled = writable &&
            RecordsGrid.SelectedItems.OfType<DataRowView>().Any();
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
        _rightClickCells = null;
        if (e.OriginalSource is not DependencyObject element) return;
        DependencyObject? current = element;
        while (current is not null)
        {
            if (current is DataGridCell cell && cell.IsSelected)
                _rightClickCells = RecordsGrid.SelectedCells.ToArray();
            if (current is DataGridRow row && row.IsSelected)
                _rightClickSelection = RecordsGrid.SelectedItems.OfType<DataRowView>()
                    .Select(view => view.Row).ToArray();
            current = VisualTreeHelper.GetParent(current);
        }
    }

    private void RecordsSelection_Changed(object sender, SelectedCellsChangedEventArgs e) =>
        UpdateActionStates();

    private void RecordsGrid_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        // When editing text inside a cell, use ordinary text editing shortcuts.
        if (e.OriginalSource is TextBox or PasswordBox) return;

        if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.C)
        {
            CopyCells_Click(sender, e);
            e.Handled = true;
        }
        else if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.V)
        {
            PasteCells_Click(sender, e);
            e.Handled = true;
        }
        else if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.D)
        {
            FillDown_Click(sender, e);
            e.Handled = true;
        }
        else if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Delete)
        {
            ClearCells_Click(sender, e);
            e.Handled = true;
        }
    }

    private sealed record SelectedGridCell(int RowIndex, int ColumnIndex, DataRow Row, DataColumn Column);

    private IReadOnlyList<SelectedGridCell> SelectedCells(object? source = null)
    {
        if (_table is null) return Array.Empty<SelectedGridCell>();

        var infos = source is MenuItem && _rightClickCells is { Count: > 0 }
            ? _rightClickCells
            : RecordsGrid.SelectedCells.ToArray();
        _rightClickCells = null;
        if (infos.Count == 0 && RecordsGrid.CurrentCell.Item is DataRowView &&
            RecordsGrid.CurrentCell.Column is not null)
            infos = new[] { RecordsGrid.CurrentCell };

        var visibleRows = _table.DefaultView.Cast<DataRowView>()
            .Select((view, index) => (view.Row, index))
            .ToDictionary(pair => pair.Row, pair => pair.index);
        var cells = new List<SelectedGridCell>();
        foreach (var info in infos)
        {
            if (info.Item is not DataRowView view ||
                info.Column?.Header is null ||
                !visibleRows.TryGetValue(view.Row, out int rowIndex) ||
                view.Row.RowState == DataRowState.Deleted)
                continue;
            string columnName = Convert.ToString(info.Column.Header) ?? "";
            if (!_table.Columns.Contains(columnName)) continue;
            cells.Add(new SelectedGridCell(rowIndex, info.Column.DisplayIndex,
                view.Row, _table.Columns[columnName]!));
        }
        return cells.DistinctBy(cell => (cell.Row, cell.Column))
            .OrderBy(cell => cell.RowIndex)
            .ThenBy(cell => cell.ColumnIndex)
            .ToArray();
    }

    private bool CommitGridEdit()
    {
        if (!RecordsGrid.CommitEdit(DataGridEditingUnit.Cell, true) ||
            !RecordsGrid.CommitEdit(DataGridEditingUnit.Row, true))
        {
            Status("Complete the active cell edit before using clipboard commands.");
            return false;
        }
        return true;
    }

    private void CopyCells_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || !CommitGridEdit()) return;
        try
        {
            var cells = SelectedCells(sender);
            string text = GridCellEditor.CopyValues(cells.Select(c =>
                new GridCellEditor.CellValue(c.RowIndex, c.ColumnIndex, c.Row[c.Column])).ToArray());
            Clipboard.SetText(text);
            Status($"Copied {cells.Count:N0} selected cells (Ctrl+C). No records were duplicated.");
        }
        catch (Exception ex) { Error("Copy cells", ex); }
    }

    private static void CheckRectangular(IReadOnlyList<SelectedGridCell> cells)
    {
        var rows = cells.Select(c => c.RowIndex).Distinct().Order().ToArray();
        var columns = cells.Select(c => c.ColumnIndex).Distinct().Order().ToArray();
        if (cells.Count != (long)rows.Length * columns.Length ||
            rows[^1] - rows[0] + 1 != rows.Length ||
            columns[^1] - columns[0] + 1 != columns.Length)
            throw new InvalidOperationException(
                "Choose a continuous rectangular destination selection for pasting multiple cells.");
    }

    private IReadOnlyList<GridCellEditor.CellPatch> CreatePatches(
        IReadOnlyList<SelectedGridCell> selected, string[][] values)
    {
        if (_table is null || selected.Count == 0)
            throw new InvalidOperationException("Select destination cells before pasting.");
        if (values.Length == 0 || values[0].Length == 0)
            throw new InvalidOperationException("Clipboard has no tabular text to paste.");

        // A single copied cell repeated into every selected destination cell
        // is the primary fill workflow, including disjoint cell selections.
        if (values.Length == 1 && values[0].Length == 1)
            return selected.Select(c => new GridCellEditor.CellPatch(c.Row, c.Column, values[0][0])).ToArray();

        CheckRectangular(selected);
        var rowIndexes = selected.Select(c => c.RowIndex).Distinct().Order().ToArray();
        var colIndexes = selected.Select(c => c.ColumnIndex).Distinct().Order().ToArray();
        int targetRows = rowIndexes.Length;
        int targetCols = colIndexes.Length;

        // With one destination cell, extend the pasted rectangle from that
        // cell, but never create additional rows or cross the visible page.
        if (selected.Count == 1)
        {
            var view = _table.DefaultView;
            var displayColumns = RecordsGrid.Columns.OrderBy(c => c.DisplayIndex).ToArray();
            var anchor = selected[0];
            if (anchor.RowIndex + values.Length > view.Count ||
                anchor.ColumnIndex + values[0].Length > displayColumns.Length)
                throw new InvalidOperationException(
                    "The pasted block exceeds the loaded rows or columns. Choose All records or select a smaller block.");

            var patches = new List<GridCellEditor.CellPatch>();
            for (int r = 0; r < values.Length; r++)
                for (int c = 0; c < values[0].Length; c++)
                {
                    string columnName = Convert.ToString(displayColumns[anchor.ColumnIndex + c].Header) ?? "";
                    if (!_table.Columns.Contains(columnName))
                        throw new InvalidOperationException("Destination column was not found.");
                    patches.Add(new GridCellEditor.CellPatch(view[anchor.RowIndex + r].Row,
                        _table.Columns[columnName]!, values[r][c]));
                }
            return patches;
        }

        // A selected rectangle can be filled by a repeating smaller block,
        // exactly as a spreadsheet tiles clipboard values into selected cells.
        if (targetRows % values.Length != 0 || targetCols % values[0].Length != 0)
            throw new InvalidOperationException(
                "The copied block does not fit the selected cells. Select a matching range or a single starting cell.");

        var rowMap = rowIndexes.Select((n, i) => (n, i)).ToDictionary(p => p.n, p => p.i);
        var colMap = colIndexes.Select((n, i) => (n, i)).ToDictionary(p => p.n, p => p.i);
        return selected.Select(cell =>
            new GridCellEditor.CellPatch(cell.Row, cell.Column,
                values[rowMap[cell.RowIndex] % values.Length]
                      [colMap[cell.ColumnIndex] % values[0].Length])).ToArray();
    }

    private void ApplyCellPatches(IReadOnlyList<GridCellEditor.CellPatch> patches, string action)
    {
        if (patches.Count > 1000 &&
            MessageBox.Show(this,
                $"{action} will update as many as {patches.Count:N0} existing cells.\n" +
                "Only the selected fields will change. Continue?",
                "Confirm bulk cell changes", MessageBoxButton.YesNo, MessageBoxImage.Question)
            != MessageBoxResult.Yes)
        {
            Status("Bulk change canceled.");
            return;
        }
        _applyingCellChanges = true;
        try
        {
            int count = GridCellEditor.Apply(patches);
            Status($"{action}: changed {count:N0} cells. Click Save Changes to write these values to the database.");
        }
        finally
        {
            _applyingCellChanges = false;
            UpdateActionStates();
        }
    }

    private void PasteCells_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || !CommitGridEdit()) return;
        try
        {
            if (!Clipboard.ContainsText())
                throw new InvalidOperationException("Copy a cell or range of cells first.");
            var selected = SelectedCells(sender);
            var pasted = GridCellEditor.ParseValues(Clipboard.GetText());
            var patches = CreatePatches(selected, pasted);
            ApplyCellPatches(patches, "Paste");
        }
        catch (Exception ex) { Error("Paste cells", ex); }
    }

    private void FillDown_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || !CommitGridEdit()) return;
        try
        {
            var selected = SelectedCells(sender);
            if (selected.Count < 2)
                throw new InvalidOperationException("Select the top cell and the cells beneath it to fill.");
            var patches = new List<GridCellEditor.CellPatch>();
            foreach (var group in selected.GroupBy(cell => cell.ColumnIndex))
            {
                var ordered = group.OrderBy(cell => cell.RowIndex).ToArray();
                if (ordered.Length < 2) continue;
                string value = GridCellEditor.ParseValues(GridCellEditor.CopyValues(new[]
                {
                    new GridCellEditor.CellValue(0, 0, ordered[0].Row[ordered[0].Column])
                }))[0][0];
                foreach (var cell in ordered.Skip(1))
                    patches.Add(new GridCellEditor.CellPatch(cell.Row, cell.Column, value));
            }
            if (patches.Count == 0)
                throw new InvalidOperationException("Select multiple cells within the same column to fill down.");
            ApplyCellPatches(patches, "Fill down");
        }
        catch (Exception ex) { Error("Fill down", ex); }
    }

    private void ClearCells_Click(object sender, RoutedEventArgs e)
    {
        if (_table is null || !CommitGridEdit()) return;
        try
        {
            var patches = SelectedCells(sender)
                .Select(c => new GridCellEditor.CellPatch(c.Row, c.Column, "")).ToArray();
            ApplyCellPatches(patches, "Clear");
        }
        catch (Exception ex) { Error("Clear cells", ex); }
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
