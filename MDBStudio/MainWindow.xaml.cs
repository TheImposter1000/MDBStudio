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
    private readonly List<TableColumnFilter> _activeColumnFilters = new();
    private int _filterSuggestionsGeneration;
    private bool _syncingFilterControls;

    private sealed record FilterChoice(TableFilterOperation Operation, string Label);

    public MainWindow(string? initialDatabase = null)
    {
        InitializeComponent();
        VersionLabel.Text = "MDB Studio " + InstalledVersion();
        Title = "MDB Studio " + InstalledVersion() + " | Database Workspace";
        if (initialDatabase is not null)
            Loaded += (_, _) => OpenDatabase(initialDatabase);
    }

    private static string InstalledVersion()
    {
        Version? version = typeof(MainWindow).Assembly.GetName().Version;
        return version is null ? "unknown" :
            "v" + version.Major + "." + version.Minor + "." + version.Build;
    }

    private static string UpdateLogPath() =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MDBStudio", "update.log");

    private void InstallationDetails_Click(object sender, RoutedEventArgs e)
    {
        string folder = Path.GetFullPath(AppContext.BaseDirectory);
        string executable = Environment.ProcessPath ?? Path.Combine(folder, "MDBStudio.exe");
        string info =
            "Running version: " + InstalledVersion() + "\n" +
            "Expected release: " + UpdateChecker.CurrentReleaseTag + "\n" +
            "Executable: " + executable + "\n" +
            "Installation folder: " + folder + "\n" +
            "Updater included: " + File.Exists(Path.Combine(folder, "MDBStudio.Updater.exe")) + "\n" +
            "Updater log: " + UpdateLogPath();
        Clipboard.SetText(info);
        MessageBox.Show(this, info + "\n\nThese details were copied to the clipboard. " +
            "If an update fails, check the installation folder and updater log.",
            "MDB Studio installation details", MessageBoxButton.OK, MessageBoxImage.Information);
    }

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
        var dialog = new OpenFileDialog
        {
            Filter = "Access databases (*.mdb;*.accdb)|*.mdb;*.accdb",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != true) return;
        OpenDatabase(dialog.FileName);
    }

    private bool OpenDatabase(string filePath)
    {
        if (!DiscardPending()) return false;
        try
        {
            AccessSession next;
            try
            {
                next = new AccessSession(filePath);
            }
            catch (DatabasePasswordRequiredException)
            {
                var prompt = new DatabasePasswordDialog(Path.GetFileName(filePath))
                {
                    Owner = this
                };
                if (prompt.ShowDialog() != true)
                {
                    Status("Database opening cancelled.");
                    return false;
                }
                // The password stays in this process and is never saved in settings or logs.
                next = new AccessSession(filePath, prompt.EnteredPassword);
            }
            var names = next.ListTables();
            _session = next;
            _table = null;
            _currentName = null;
            _activeColumnFilters.Clear();
            ResetFilterControls();
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
            ExportQueryExcelButton.IsEnabled = false;
            ExportQueryCsvButton.IsEnabled = false;
            ImportExcelButton.IsEnabled = true;
            SchemaButton.IsEnabled = true;
            NewTableButton.IsEnabled = true;
            SidebarCreateTableButton.IsEnabled = true;
            SidebarDesignViewButton.IsEnabled = false;
            EmptyDatabasePanel.Visibility = names.Count == 0
                ? Visibility.Visible : Visibility.Collapsed;
            FileLabel.Text = filePath;
            DatabaseTitle.Text = Path.GetFileName(filePath);
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
            Status(names.Count == 0
                ? "Opened " + Path.GetFileName(filePath) +
                  ". No tables yet. Choose Create table to begin."
                : "Opened " + Path.GetFileName(filePath) + ". " +
                  names.Count + " tables found.");
            return true;
        }
        catch (Exception ex)
        {
            _suppressSelection = false;
            Error("Could not open database", ex);
            return false;
        }
    }

    private void CreateDatabase_Click(object sender, RoutedEventArgs e)
    {
        if (!DiscardPending()) return;
        var dialog = new SaveFileDialog
        {
            Title = "Create a blank Access database",
            Filter = "Access MDB database (*.mdb)|*.mdb|Access ACCDB database (*.accdb)|*.accdb",
            FileName = "NewDatabase.mdb",
            DefaultExt = ".mdb",
            AddExtension = true,
            OverwritePrompt = false
        };
        if (dialog.ShowDialog(this) != true) return;
        try
        {
            string destination = AccessDatabaseCreator.Create(dialog.FileName);
            if (!OpenDatabase(destination)) return;
            WorkspaceTabs.SelectedIndex = 0;
            // Access-style creation: immediately design the first usable table
            // instead of leaving the user in an empty, inaccessible workspace.
            if (!CreateTableInCurrentDatabase())
                Status("Created " + Path.GetFileName(destination) +
                       ". The database has no tables yet. Choose Create table whenever you are ready.");
        }
        catch (Exception ex) { Error("Create database failed", ex); }
    }

    private void CreateTable_Click(object sender, RoutedEventArgs e)
    {
        CreateTableInCurrentDatabase();
    }

    private bool CreateTableInCurrentDatabase()
    {
        if (_session is null)
        {
            MessageBox.Show(this, "Create or open a database first.",
                "Create table", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }
        if (!CommitGridEdit()) return false;
        if (_table?.GetChanges() is not null)
        {
            MessageBox.Show(this,
                "Save or discard existing record changes before creating a table. " +
                "The current grid cannot be reloaded while edits are pending.",
                "Unsaved record edits", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        try
        {
            string proposed = DatabaseWorkspaceSetup.SuggestTableName(_session.ListTables());
            var designer = new TableDesignerWindow(suggestedTableName: proposed) { Owner = this };
            if (designer.ShowDialog() != true) return false;

            string name = designer.TableName;
            _session.CreateTable(name, designer.Definition); // Creates a backup first.
            var names = _session.ListTables();
            _suppressSelection = true;
            _allTables = names;
            TableSearchBox.Clear();
            TableList.ItemsSource = names;
            TableList.SelectedItem = name;
            _suppressSelection = false;
            _table = null;
            _currentName = null;
            _activeColumnFilters.Clear();
            ResetFilterControls();
            EmptyDatabasePanel.Visibility = Visibility.Collapsed;
            WorkspaceTabs.SelectedIndex = 0;
            if (!LoadTable(name, 0))
                Status("Created " + name +
                       ". Choose the table in the left pane to view its records.");
            else
                Status("Created table " + name +
                       ". Add records in Datasheet View and select Save changes when finished.");
            return true;
        }
        catch (Exception ex)
        {
            _suppressSelection = false;
            Error("Create table failed", ex);
            return false;
        }
    }

    private void AssociateFiles_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string application = Environment.ProcessPath ??
                Path.Combine(AppContext.BaseDirectory, "MDBStudio.exe");
            WindowsFileAssociation.RegisterForCurrentUser(application);
            var choice = MessageBox.Show(this,
                "MDB Studio has been added as an available application for .mdb and .accdb files. " +
                "It has not changed your current Windows default app.\n\n" +
                "Open Windows Default Apps settings now? Choose MDB Studio for .mdb and .accdb. " +
                "Afterward, double-clicking a database in File Explorer will open it in MDB Studio.",
                "Open database files with MDB Studio",
                MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (choice == MessageBoxResult.Yes)
                Process.Start(new ProcessStartInfo(WindowsFileAssociation.DefaultAppsSettingsUri)
                {
                    UseShellExecute = true
                });
        }
        catch (Exception ex) { Error("Could not register file associations", ex); }
    }

    private void ManageSchema_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null) return;
        if (!CommitGridEdit()) return;
        if (_table?.GetChanges() is not null)
        {
            MessageBox.Show(this,
                "Save or discard pending record edits before changing the database structure. " +
                "This prevents changes from being lost when a table reloads.",
                "Unsaved record edits", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var manager = new SchemaManagerWindow(_session, _currentName) { Owner = this };
        manager.ShowDialog();
        if (!manager.StructureChanged) return;
        try
        {
            var tables = _session.ListTables();
            _suppressSelection = true;
            _allTables = tables;
            EmptyDatabasePanel.Visibility = tables.Count == 0
                ? Visibility.Visible : Visibility.Collapsed;
            TableSearchBox.Clear();
            TableList.ItemsSource = tables;
            TableList.SelectedItem = null;
            _suppressSelection = false;
            _table = null;
            _currentName = null;
            _activeColumnFilters.Clear();
            ResetFilterControls();
            RecordsGrid.ItemsSource = null;
            RecordSearchBox.Clear();
            TableLabel.Text = "Choose a table";
            KeyLabel.Text = "";
            _totalRecords = 0;
            _pageIndex = 0;
            UpdatePageControls(0);
            if (manager.SelectedTable is string selected &&
                tables.Contains(selected, StringComparer.OrdinalIgnoreCase))
            {
                _suppressSelection = true;
                TableList.SelectedItem = tables.First(n =>
                    n.Equals(selected, StringComparison.OrdinalIgnoreCase));
                _suppressSelection = false;
                LoadTable(selected, 0);
            }
            else
                UpdateActionStates();
            Status("Database structure updated. A backup was created before each schema change.");
        }
        catch (Exception ex)
        {
            _suppressSelection = false;
            Error("Could not refresh database tables", ex);
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

    private bool LoadTable(string name, int requestedPage)
    {
        try
        {
            bool switchingTable = !string.Equals(_currentName, name, StringComparison.OrdinalIgnoreCase);
            var page = _session!.LoadPage(name, requestedPage, _pageSize,
                switchingTable ? Array.Empty<TableColumnFilter>() : _activeColumnFilters);
            var table = page.Rows;
            _table = table;
            _currentName = name;
            if (switchingTable)
            {
                _activeColumnFilters.Clear();
                PopulateFilterColumns(table);
            }
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
            KeyLabel.Text = writable ? "Edit enabled (primary key detected)" :
                "Cell editing disabled (no primary key); record deletion is available.";
            UpdatePageControls(table.Rows.Count);
            RefreshFilterSummary();
            Status(_activeColumnFilters.Count == 0
                ? $"Loaded {table.Rows.Count:N0} rows from {name}. Total: {_totalRecords:N0}."
                : $"Filtered {name}: {_totalRecords:N0} matching records across the entire database.");
            return true;
        }
        catch (Exception ex)
        {
            Error("Could not load table", ex);
            return false;
        }
    }

    private static FilterChoice[] ChoicesFor(DataColumn column)
    {
        var choices = new List<FilterChoice>();
        if (column.DataType == typeof(string))
        {
            choices.Add(new(TableFilterOperation.Contains, "Contains"));
            choices.Add(new(TableFilterOperation.StartsWith, "Starts with"));
            choices.Add(new(TableFilterOperation.EndsWith, "Ends with"));
        }
        choices.Add(new(TableFilterOperation.Equals, "Equals"));
        choices.Add(new(TableFilterOperation.NotEquals, "Does not equal"));
        if (TableFilters.Supports(column, TableFilterOperation.GreaterThan))
        {
            choices.Add(new(TableFilterOperation.GreaterThan, "Greater than"));
            choices.Add(new(TableFilterOperation.LessThan, "Less than"));
        }
        choices.Add(new(TableFilterOperation.IsBlank, "Is blank"));
        choices.Add(new(TableFilterOperation.IsNotBlank, "Is not blank"));
        return choices.ToArray();
    }

    private void ResetFilterControls()
    {
        _filterSuggestionsGeneration++;
        _syncingFilterControls = true;
        try
        {
            FilterColumnCombo.ItemsSource = null;
            FilterColumnCombo.SelectedIndex = -1;
            FilterColumnCombo.IsEnabled = false;
            FilterOperationCombo.ItemsSource = null;
            FilterOperationCombo.IsEnabled = false;
            FilterValueCombo.ItemsSource = null;
            FilterValueCombo.Text = "";
            FilterValueCombo.IsEnabled = false;
            ApplyColumnFilterButton.IsEnabled = false;
            RemoveColumnFilterButton.IsEnabled = false;
            ClearColumnFiltersButton.IsEnabled = false;
            ActiveFiltersLabel.Text = "No active filters. Click a column header to choose a filter.";
        }
        finally { _syncingFilterControls = false; }
    }

    private void PopulateFilterColumns(DataTable table)
    {
        ResetFilterControls();
        _syncingFilterControls = true;
        try
        {
            FilterColumnCombo.ItemsSource = table.Columns.Cast<DataColumn>()
                .Where(TableFilters.CanFilter)
                .Select(column => column.ColumnName)
                .ToArray();
            FilterColumnCombo.IsEnabled = FilterColumnCombo.Items.Count > 0;
        }
        finally { _syncingFilterControls = false; }
    }

    private void RefreshFilterSummary()
    {
        if (ActiveFiltersLabel is null) return;
        ActiveFiltersLabel.Text = _activeColumnFilters.Count == 0
            ? "No active filters. Click any column header or select a header above."
            : "Active across the whole table: " + string.Join("  |  ",
                _activeColumnFilters.Select(f => f.ColumnName + " " +
                    f.Operation.ToString() + (f.Operation is TableFilterOperation.IsBlank or
                        TableFilterOperation.IsNotBlank ? "" : " " + f.Value)));
        ClearColumnFiltersButton.IsEnabled = _activeColumnFilters.Count > 0;
        RemoveColumnFilterButton.IsEnabled = FilterColumnCombo.SelectedItem is string name &&
            _activeColumnFilters.Any(f => string.Equals(f.ColumnName, name,
                StringComparison.OrdinalIgnoreCase));
    }

    private void FilterColumn_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingFilterControls || _table is null) return;
        _filterSuggestionsGeneration++;
        _syncingFilterControls = true;
        try
        {
            FilterValueCombo.ItemsSource = null;
            FilterValueCombo.Text = "";
            if (FilterColumnCombo.SelectedItem is not string name ||
                !_table.Columns.Contains(name))
            {
                FilterOperationCombo.ItemsSource = null;
                FilterOperationCombo.IsEnabled = false;
                FilterValueCombo.IsEnabled = false;
                ApplyColumnFilterButton.IsEnabled = false;
                return;
            }

            var column = _table.Columns[name]!;
            var choices = ChoicesFor(column);
            var current = _activeColumnFilters.FirstOrDefault(f =>
                string.Equals(f.ColumnName, name, StringComparison.OrdinalIgnoreCase));
            FilterOperationCombo.ItemsSource = choices;
            FilterOperationCombo.IsEnabled = true;
            FilterOperationCombo.SelectedItem = choices.FirstOrDefault(c =>
                c.Operation == current?.Operation) ?? choices[0];
            FilterValueCombo.Text = current?.Value ?? "";
            FilterValueCombo.IsEnabled =
                ((FilterChoice)FilterOperationCombo.SelectedItem).Operation is not
                    (TableFilterOperation.IsBlank or TableFilterOperation.IsNotBlank);
            ApplyColumnFilterButton.IsEnabled = true;
        }
        finally
        {
            _syncingFilterControls = false;
            RefreshFilterSummary();
        }

        if (FilterColumnCombo.SelectedItem is string selected &&
            _session is AccessSession session && _currentName is string tableName)
        {
            int version = _filterSuggestionsGeneration;
            _ = LoadFilterSuggestionsAsync(session, tableName, selected, version);
        }
    }

    private async Task LoadFilterSuggestionsAsync(
        AccessSession session, string tableName, string column, int generation)
    {
        try
        {
            var options = await Task.Run(() => session.DistinctColumnValues(tableName, column));
            if (generation != _filterSuggestionsGeneration || _session != session ||
                _currentName != tableName ||
                !string.Equals(FilterColumnCombo.SelectedItem as string, column, StringComparison.Ordinal))
                return;
            string typed = FilterValueCombo.Text;
            FilterValueCombo.ItemsSource = options;
            FilterValueCombo.Text = typed;
        }
        catch (Exception)
        {
            // Large/legacy Access columns may not support DISTINCT.
            // Free-text value entry remains available regardless.
        }
    }

    private void FilterOperation_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_syncingFilterControls) return;
        FilterValueCombo.IsEnabled = FilterOperationCombo.SelectedItem is FilterChoice choice &&
            choice.Operation is not (TableFilterOperation.IsBlank or TableFilterOperation.IsNotBlank);
    }

    private void RecordsGrid_Sorting(object sender, DataGridSortingEventArgs e)
    {
        // Header click opens the filter controls rather than sorting one page only.
        e.Handled = true;
        var name = Convert.ToString(e.Column.Header);
        if (!string.IsNullOrEmpty(name) && FilterColumnCombo.Items.Contains(name))
        {
            FilterColumnCombo.SelectedItem = name;
            if (FilterValueCombo.IsEnabled) FilterValueCombo.Focus();
        }
    }

    private void ApplyColumnFilter_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null || _currentName is null || _table is null ||
            FilterColumnCombo.SelectedItem is not string column ||
            FilterOperationCombo.SelectedItem is not FilterChoice choice) return;

        var next = _activeColumnFilters
            .Where(f => !string.Equals(f.ColumnName, column, StringComparison.OrdinalIgnoreCase))
            .ToList();
        next.Add(new TableColumnFilter(column, choice.Operation, FilterValueCombo.Text));
        try { TableFilters.Compile(_table.Columns, next); }
        catch (Exception ex) when (ex is FormatException or InvalidOperationException)
        {
            Error("Invalid column filter", ex);
            return;
        }

        if (!DiscardPending()) return;
        var prior = _activeColumnFilters.ToArray();
        _activeColumnFilters.Clear();
        _activeColumnFilters.AddRange(next);
        if (!LoadTable(_currentName, 0))
        {
            _activeColumnFilters.Clear();
            _activeColumnFilters.AddRange(prior);
        }
        RefreshFilterSummary();
    }

    private void RemoveColumnFilter_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || FilterColumnCombo.SelectedItem is not string column ||
            !_activeColumnFilters.Any(f =>
                string.Equals(f.ColumnName, column, StringComparison.OrdinalIgnoreCase)) ||
            !DiscardPending()) return;

        var prior = _activeColumnFilters.ToArray();
        _activeColumnFilters.RemoveAll(f => string.Equals(f.ColumnName, column,
            StringComparison.OrdinalIgnoreCase));
        if (!LoadTable(_currentName, 0))
        {
            _activeColumnFilters.Clear();
            _activeColumnFilters.AddRange(prior);
        }
        RefreshFilterSummary();
    }

    private void ClearColumnFilters_Click(object sender, RoutedEventArgs e)
    {
        if (_currentName is null || _activeColumnFilters.Count == 0 || !DiscardPending()) return;
        var prior = _activeColumnFilters.ToArray();
        _activeColumnFilters.Clear();
        if (!LoadTable(_currentName, 0))
            _activeColumnFilters.AddRange(prior);
        RefreshFilterSummary();
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

    private void VisualQueryDesigner_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null)
        {
            MessageBox.Show(this, "Open an MDB or ACCDB database first.",
                "Visual query designer", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var designer = new VisualQueryDesignerWindow(_session, _currentName)
        {
            Owner = this
        };
        if (designer.ShowDialog() == true && designer.ResultSql is string sql)
        {
            SqlInput.Text = sql;
            WorkspaceTabs.SelectedIndex = 1;
            Query_Click(sender, e);
        }
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
            ExportQueryExcelButton.IsEnabled = true;
            ExportQueryCsvButton.IsEnabled = true;
            Status("Query returned " + results.Rows.Count.ToString("N0") + " rows.");
        }
        catch (Exception ex) { Error("Query failed", ex); }
    }

    private async void ExportQueryExcel_Click(object sender, RoutedEventArgs e)
    {
        if (QueryGrid.ItemsSource is not DataView view)
        {
            MessageBox.Show(this, "Run a SELECT query first.", "Excel export");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "Excel workbooks (*.xlsx)|*.xlsx",
            FileName = "MDBStudio_Query_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx",
            DefaultExt = ".xlsx",
            AddExtension = true,
            OverwritePrompt = true
        };
        if (dialog.ShowDialog(this) != true) return;

        ExportQueryExcelButton.IsEnabled = false;
        try
        {
            var results = view.Table.Copy();
            Status($"Exporting {results.Rows.Count:N0} query results to Excel...");
            await Task.Run(() => ExcelTransfer.ExportQuery(results, dialog.FileName));
            Status($"Exported {results.Rows.Count:N0} query results to {dialog.FileName}.");
            MessageBox.Show(this, "Excel query export completed successfully.",
                "Export to Excel", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { Error("Excel export failed", ex); }
        finally { ExportQueryExcelButton.IsEnabled = true; }
    }

    private async void ExportQueryCsv_Click(object sender, RoutedEventArgs e)
    {
        if (QueryGrid.ItemsSource is not DataView view)
        {
            MessageBox.Show(this, "Run a SELECT query first.", "CSV export");
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv",
            FileName = "MDBStudio_Query_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".csv",
            DefaultExt = ".csv",
            AddExtension = true,
            OverwritePrompt = true
        };
        if (dialog.ShowDialog(this) != true) return;

        ExportQueryCsvButton.IsEnabled = false;
        try
        {
            var results = view.Table.Copy();
            Status($"Exporting {results.Rows.Count:N0} query results to CSV...");
            await Task.Run(() => CsvTransfer.ExportQuery(results, dialog.FileName));
            Status($"Exported {results.Rows.Count:N0} query results to {dialog.FileName}.");
            MessageBox.Show(this, "CSV query export completed successfully.",
                "Export to CSV", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex) { Error("CSV export failed", ex); }
        finally { ExportQueryCsvButton.IsEnabled = true; }
    }

    private async void ImportExcel_Click(object sender, RoutedEventArgs e)
    {
        if (_session is null) return;
        if (!CommitGridEdit()) return;
        if (_table?.GetChanges() is not null)
        {
            MessageBox.Show(this,
                "Save or discard your pending database edits before importing Excel or CSV. " +
                "The import commits separately and must not overwrite your unsaved work.",
                "Unsaved records", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var file = new OpenFileDialog
        {
            Filter = "Excel and CSV files (*.xlsx;*.csv)|*.xlsx;*.csv|Excel workbooks (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv",
            CheckFileExists = true,
            Multiselect = false
        };
        if (file.ShowDialog(this) != true) return;
        try
        {
            var currentSession = _session!;
            var wizard = new ExcelImportDialog(file.FileName, _allTables) { Owner = this };
            if (wizard.ShowDialog() != true || wizard.SelectedSheet is null) return;

            var data = wizard.SelectedSheet;
            var target = wizard.Destination;
            var create = wizard.CreateNewTable;
            ImportExcelButton.IsEnabled = false;
            Status($"Importing {data.Rows.Count:N0} spreadsheet records into {target}; backing up MDB...");
            try
            {
                // File and database work happen off the UI thread.
                int imported = await Task.Run(() => currentSession.ImportSpreadsheet(data, target, create));
                var names = currentSession.ListTables();
                _allTables = names;
                _suppressSelection = true;
                TableSearchBox.Clear();
                TableList.ItemsSource = names;
                TableList.SelectedItem = target;
                _suppressSelection = false;
                _activeColumnFilters.Clear();
                _currentName = null; // Reload the destination with fresh filter controls.
                LoadTable(target, 0);
                Status($"Import complete: {imported:N0} records into {target}. MDB backup created.");
                MessageBox.Show(this,
                    $"Imported {imported:N0} records successfully into '{target}'.\n\n" +
                    "The import was committed separately and a backup of your MDB was created.",
                    "Excel import complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            finally { ImportExcelButton.IsEnabled = true; }
        }
        catch (Exception ex)
        {
            Error("Spreadsheet import failed", ex);
            Status("Import was not completed. Check the error and your MDB backup.");
        }
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
            Status("Installed: " + InstalledVersion() + ", latest published: " + result.LatestTag + ".");
            if (!result.UpdateAvailable)
            {
                Status("MDB Studio " + InstalledVersion() + " is up to date.");
                MessageBox.Show(this, "Installed executable: " + InstalledVersion() + "\nLatest GitHub release: " + result.LatestTag + "\n\n" +
                    "Installation folder: " + Path.GetFullPath(AppContext.BaseDirectory),
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
                $"MDB Studio {result.LatestTag} is available.\n\nInstalled executable: {InstalledVersion()}\n" +
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
            // Verify the portable installation folder is writable BEFORE
            // exiting; protected folders cannot be updated without elevation.
            string probe = Path.Combine(AppContext.BaseDirectory,
                ".mdbstudio-write-test-" + Guid.NewGuid().ToString("N"));
            try
            {
                using (var stream = new FileStream(probe, FileMode.CreateNew,
                    FileAccess.Write, FileShare.None)) stream.WriteByte(1);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException(
                    "MDB Studio cannot update files in its current installation folder. " +
                    "Move the portable app to a folder your Windows account can write to " +
                    "and retry. Current folder: " + Path.GetFullPath(AppContext.BaseDirectory), ex);
            }
            finally { if (File.Exists(probe)) File.Delete(probe); }
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
        SidebarDesignViewButton.IsEnabled = _session is not null && _currentName is not null;
        int cellCount = RecordsGrid.SelectedCells.Count;
        CopyButton.IsEnabled = open && cellCount > 0;
        PasteButton.IsEnabled = writable && cellCount > 0;
        FillDownButton.IsEnabled = writable && cellCount > 1;
        // Cell selection is the default in v1.0.3, so deleting records must
        // also be available when the user selects cells rather than row headers.
        DeleteButton.IsEnabled = open &&
            (RecordsGrid.SelectedItems.OfType<DataRowView>().Any() ||
             RecordsGrid.SelectedCells.Any(cell => cell.Item is DataRowView));
        // Keyless tables remain read-only for edits, but staged deletions must
        // still enable Save Changes so the user can commit them.
        SaveButton.IsEnabled = writable ||
            (open && _table!.Rows.Cast<DataRow>().Any(row => row.RowState == DataRowState.Deleted));
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
        // Row-header selections take priority; otherwise collect the rows
        // belonging to selected cells. Multiple cells in one row count once.
        // Right-clicking a selected cell retains the original selection.
        var fromContext = source is MenuItem;
        var rows = fromContext && _rightClickSelection is { Count: > 0 } context
            ? context
            : RecordsGrid.SelectedItems.OfType<DataRowView>()
                .Select(view => view.Row).ToArray();
        if (rows.Count == 0)
        {
            var cells = fromContext && _rightClickCells is { Count: > 0 }
                ? _rightClickCells
                : RecordsGrid.SelectedCells.ToArray();
            rows = cells.Where(cell => cell.Item is DataRowView)
                .Select(cell => ((DataRowView)cell.Item).Row).ToArray();
        }
        _rightClickSelection = null;
        _rightClickCells = null;
        return RecordDeletion.EligibleRows(rows);
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
        else if ((Keyboard.Modifiers & ModifierKeys.Control) != 0 && e.Key == Key.Delete)
        {
            // Ctrl+Delete explicitly deletes records even when cells are selected.
            DeleteRecords_Click(sender, e);
            e.Handled = true;
        }
        else if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Delete)
        {
            // Delete on selected row headers removes records. With ordinary
            // cell selection it continues to clear just those cell values.
            if (RecordsGrid.SelectedItems.OfType<DataRowView>().Any())
                DeleteRecords_Click(sender, e);
            else
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
        if (_table is null) return;
        if (!CommitGridEdit()) return;
        try
        {
            var rows = SelectedRows(sender);
            if (rows.Count == 0)
            {
                Status("Select any cell in the records to delete, or select their row headers.");
                return;
            }
            if (MessageBox.Show(this,
                $"Delete {rows.Count:N0} selected record(s)?\n\n" +
                "Deletion remains pending until you click Save Changes. MDB Studio " +
                "backs up the file before saving. On tables without a primary key, " +
                "MDB Studio checks matching records to prevent accidental deletion " +
                "of unselected duplicates.",
                "Confirm record deletion",
                MessageBoxButton.YesNo, MessageBoxImage.Warning,
                MessageBoxResult.No) != MessageBoxResult.Yes)
            {
                Status("Record deletion canceled.");
                return;
            }

            int count = RecordDeletion.Stage(_table, rows);
            RecordsGrid.SelectedCells.Clear();
            Status($"{count:N0} records marked for deletion. Select Save Changes to commit.");
            UpdateActionStates();
        }
        catch (Exception ex) { Error("Delete records", ex); }
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
