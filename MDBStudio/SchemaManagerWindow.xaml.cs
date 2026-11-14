using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MDBStudio;

/// <summary>
/// Schema operations are separate from Save Changes; Access DDL is applied
/// immediately after a fresh on-disk backup. The user can always cancel.
/// </summary>
public partial class SchemaManagerWindow : Window
{
    private readonly AccessSession _session;
    private bool _updating;
    private bool _ready;
    private List<SchemaColumnInfo> _fieldOrder = new();
    private string[] _originalFieldOrder = Array.Empty<string>();
    private Point _dragStart;
    private SchemaColumnInfo? _dragSource;

    public string? SelectedTable => TablesCombo.SelectedItem as string;
    public bool StructureChanged { get; private set; }

    public SchemaManagerWindow(AccessSession session, string? initialTable = null)
    {
        InitializeComponent();
        _session = session;
        _ready = true;
        RefreshTables(initialTable);
    }

    private void RefreshTables(string? preferred = null)
    {
        try
        {
            string? oldParent = ParentTableCombo.SelectedItem as string;
            string? table = preferred ?? SelectedTable;
            _updating = true;
            var names = _session.ListTables();
            TablesCombo.ItemsSource = names;
            ParentTableCombo.ItemsSource = names;
            if (table is not null && names.Contains(table, StringComparer.OrdinalIgnoreCase))
                TablesCombo.SelectedItem = names.First(n => n.Equals(table, StringComparison.OrdinalIgnoreCase));
            else if (names.Count > 0) TablesCombo.SelectedIndex = 0;

            if (oldParent is not null &&
                names.Contains(oldParent, StringComparer.OrdinalIgnoreCase))
                ParentTableCombo.SelectedItem = names.First(n =>
                    n.Equals(oldParent, StringComparison.OrdinalIgnoreCase));
            else if (names.Count > 0) ParentTableCombo.SelectedIndex = 0;
        }
        catch (Exception ex) { ShowError("Unable to load table names", ex); }
        finally { _updating = false; }
        DisplayTable();
    }

    private void DisplayTable()
    {
        if (!_ready || _updating) return;
        if (SelectedTable is not string table)
        {
            FieldsGrid.ItemsSource = null;
            _fieldOrder.Clear();
            _originalFieldOrder = Array.Empty<string>();
            IndexColumnCombo.ItemsSource = null;
            ChildFieldCombo.ItemsSource = null;
            IndexesCombo.ItemsSource = null;
            SetTableActions(false);
            PopulateParentFields();
            return;
        }

        try
        {
            var info = _session.InspectTable(table);
            var byName = info.Columns.ToDictionary(c => c.Name, StringComparer.OrdinalIgnoreCase);
            var currentOrder = _session.ReadDesignFieldOrder(table);
            _fieldOrder = currentOrder.Select(name => byName[name]).ToList();
            _originalFieldOrder = currentOrder.ToArray();
            FieldsGrid.ItemsSource = _fieldOrder;
            UpdateFieldOrderButtons();
            string[] columns = info.Columns.Select(c => c.Name).ToArray();
            IndexColumnCombo.ItemsSource = columns;
            ChildFieldCombo.ItemsSource = columns;
            if (columns.Length > 0)
            {
                IndexColumnCombo.SelectedIndex = 0;
                ChildFieldCombo.SelectedIndex = 0;
            }
            IndexesCombo.ItemsSource = info.IndexNames;
            if (info.IndexNames.Count > 0) IndexesCombo.SelectedIndex = 0;
            SetTableActions(true);
            IndexNameBox.Text = "IX_" + table[..Math.Min(table.Length, 61)];
            RelationNameBox.Text = "FK_" + table[..Math.Min(table.Length, 61)];
            PopulateParentFields();
        }
        catch (Exception ex) { ShowError("Could not read table design", ex); }
    }

    private void SetTableActions(bool enabled)
    {
        AddFieldButton.IsEnabled = enabled;
        MoveFieldUpButton.IsEnabled = enabled;
        MoveFieldDownButton.IsEnabled = enabled;
        MatchDatasheetOrderButton.IsEnabled = enabled;
        SaveFieldOrderButton.IsEnabled = false;
        ChangeFieldButton.IsEnabled = enabled;
        EditTableDesignButton.IsEnabled = enabled;
        DropFieldButton.IsEnabled = enabled;
        DeleteTableButton.IsEnabled = enabled;
        CreateIndexButton.IsEnabled = enabled;
        DropIndexButton.IsEnabled = enabled && IndexesCombo.Items.Count > 0;
        CreateRelationButton.IsEnabled = enabled;
    }

    private void PopulateParentFields()
    {
        if (!_ready || _updating) return;
        if (ParentTableCombo.SelectedItem is not string table)
        {
            ParentFieldCombo.ItemsSource = null;
            return;
        }
        try
        {
            var info = _session.InspectTable(table);
            ParentFieldCombo.ItemsSource = info.Columns
                .Where(c => c.PrimaryKey).Select(c => c.Name).ToArray();
            if (ParentFieldCombo.Items.Count > 0) ParentFieldCombo.SelectedIndex = 0;
        }
        catch (Exception ex) { ShowError("Cannot read referenced table", ex); }
    }

    private void UpdateFieldOrderButtons()
    {
        bool hasTable = SelectedTable is not null && _fieldOrder.Count > 0;
        MoveFieldUpButton.IsEnabled = hasTable && _fieldOrder.Count > 1;
        MoveFieldDownButton.IsEnabled = hasTable && _fieldOrder.Count > 1;
        MatchDatasheetOrderButton.IsEnabled = hasTable;
        SaveFieldOrderButton.IsEnabled = hasTable &&
            !AccessFieldOrder.Same(_originalFieldOrder,
                _fieldOrder.Select(field => field.Name).ToArray());
    }

    private void RedrawFields(SchemaColumnInfo? selected)
    {
        FieldsGrid.ItemsSource = null;
        FieldsGrid.ItemsSource = _fieldOrder;
        if (selected is not null)
        {
            FieldsGrid.SelectedItem = selected;
            FieldsGrid.ScrollIntoView(selected);
        }
        UpdateFieldOrderButtons();
        if (SaveFieldOrderButton.IsEnabled)
            ResultLabel.Text = "Field order changed. Select Save field order to write it into the MDB.";
    }

    private void MoveSelectedField(int direction)
    {
        if (FieldsGrid.SelectedItem is not SchemaColumnInfo selected) return;
        int index = _fieldOrder.IndexOf(selected);
        int destination = index + direction;
        if (index < 0 || destination < 0 || destination >= _fieldOrder.Count) return;
        _fieldOrder.RemoveAt(index);
        _fieldOrder.Insert(destination, selected);
        RedrawFields(selected);
    }

    private void MoveFieldUp_Click(object sender, RoutedEventArgs e) => MoveSelectedField(-1);
    private void MoveFieldDown_Click(object sender, RoutedEventArgs e) => MoveSelectedField(1);

    private void MatchDatasheetOrder_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table) return;
        var preferred = TableColumnLayout.Load(_session.FilePath, table);
        if (preferred.Count == 0)
        {
            ResultLabel.Text = "No custom datasheet column arrangement saved for " + table + ".";
            return;
        }

        var desired = TableColumnLayout.Apply(
            _fieldOrder.Select(f => f.Name).ToArray(), preferred);
        var lookup = _fieldOrder.ToDictionary(f => f.Name, StringComparer.OrdinalIgnoreCase);
        _fieldOrder = desired.Select(name => lookup[name]).ToList();
        RedrawFields(null);
        ResultLabel.Text = SaveFieldOrderButton.IsEnabled
            ? "Datasheet arrangement loaded. Select Save field order to apply it to the database."
            : "The database field order already matches your datasheet arrangement.";
    }

    private void SaveFieldOrder_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table || !SaveFieldOrderButton.IsEnabled) return;
        string[] names = _fieldOrder.Select(f => f.Name).ToArray();
        Apply(() =>
        {
            // DAO updates Access's real TableDef.Fields.OrdinalPosition,
            // even when the table already contains records.
            _session.ReorderTableFields(table, names);
            // The datasheet should now reflect the persisted schema order.
            TableColumnLayout.Save(_session.FilePath, table, names);
        }, "Saved actual field order for " + table + ".", table);
    }

    private static DataGridRow? FieldRowFromSource(object source)
    {
        DependencyObject? current = source as DependencyObject;
        while (current is not null)
        {
            if (current is DataGridRow row) return row;
            current = current is Visual ? VisualTreeHelper.GetParent(current) : null;
        }
        return null;
    }

    private void FieldsGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _dragStart = e.GetPosition(FieldsGrid);
        _dragSource = FieldRowFromSource(e.OriginalSource)?.Item as SchemaColumnInfo;
    }

    private void FieldsGrid_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (_dragSource is null || e.LeftButton != MouseButtonState.Pressed) return;
        Point pointer = e.GetPosition(FieldsGrid);
        if (Math.Abs(pointer.X - _dragStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(pointer.Y - _dragStart.Y) < SystemParameters.MinimumVerticalDragDistance)
            return;

        SchemaColumnInfo source = _dragSource;
        _dragSource = null;
        DragDrop.DoDragDrop(FieldsGrid,
            new DataObject(typeof(SchemaColumnInfo), source), DragDropEffects.Move);
    }

    private void FieldsGrid_DragOver(object sender, DragEventArgs e)
    {
        e.Effects = e.Data.GetDataPresent(typeof(SchemaColumnInfo))
            ? DragDropEffects.Move : DragDropEffects.None;
        e.Handled = true;
    }

    private void FieldsGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetData(typeof(SchemaColumnInfo)) is not SchemaColumnInfo from ||
            FieldRowFromSource(e.OriginalSource)?.Item is not SchemaColumnInfo to)
            return;

        int source = _fieldOrder.IndexOf(from);
        int destination = _fieldOrder.IndexOf(to);
        if (source >= 0 && destination >= 0 && source != destination)
        {
            _fieldOrder.RemoveAt(source);
            _fieldOrder.Insert(destination, from);
            RedrawFields(from);
        }
        e.Handled = true;
    }

    private void TableChanged(object sender, SelectionChangedEventArgs e) => DisplayTable();
    private void ParentTableChanged(object sender, SelectionChangedEventArgs e) => PopulateParentFields();
    private void Refresh_Click(object sender, RoutedEventArgs e) => RefreshTables();
    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private void Apply(Action operation, string message, string? select = null)
    {
        try
        {
            operation();
            StructureChanged = true;
            RefreshTables(select);
            ResultLabel.Text = message + " A database backup was created.";
        }
        catch (Exception ex) { ShowError("Database structure change failed", ex); }
    }

    private void CreateTable_Click(object sender, RoutedEventArgs e)
    {
        var designer = new TableDesignerWindow(
            suggestedTableName: DatabaseWorkspaceSetup.SuggestTableName(_session.ListTables()))
            { Owner = this };
        if (designer.ShowDialog() != true) return;
        string table = designer.TableName;
        var fields = designer.Definition;
        Apply(() => _session.CreateTable(table, fields), "Created table " + table + ".", table);
    }

    private void AddField_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table) return;
        var designer = new TableDesignerWindow(table) { Owner = this };
        if (designer.ShowDialog() != true) return;
        var field = designer.Definition[0];
        Apply(() => _session.AddField(table, field),
            "Added " + field.Name + " to " + table + ".", table);
    }

    private void EditTableDesign_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table) return;
        try
        {
            // For empty tables we can safely rebuild the whole design on a
            // staging database. A populated table must never be dropped.
            long rows = _session.TableRecordCount(table);
            if (rows != 0)
            {
                MessageBox.Show(this,
                    "This table contains " + rows.ToString("N0") + " records. " +
                    "Full table redesign is available only for empty tables, " +
                    "because changing an AutoNumber or primary key can break existing records and relationships.\n\n" +
                    "You can still add fields, change ordinary field types, and create indexes. " +
                    "Your table has not been modified.",
                    "Preserving existing records", MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var schema = _session.InspectTable(table);
            var designer = new TableDesignerWindow(
                existingTable: table, redesignFields: schema.Columns)
            {
                Owner = this
            };
            if (designer.ShowDialog() != true) return;

            EmptyTableRedesignPlan plan =
                _session.PlanEmptyTableRedesign(table, designer.Definition);
            Apply(() => _session.RedesignEmptyTable(plan),
                "Updated table design for " + table + ".", table);
        }
        catch (Exception ex)
        {
            ShowError("Cannot redesign this table", ex);
        }
    }

    private void ChangeField_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table ||
            FieldsGrid.SelectedItem is not SchemaColumnInfo field)
        {
            MessageBox.Show(this, "Select a field row to change its data type.",
                "Select a field", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        // An empty table has no records to convert. Use its complete Design
        // View instead of ALTER COLUMN, which ACE can reject even on an empty
        // column with a "Data type mismatch in criteria expression" error.
        try
        {
            if (_session.TableRecordCount(table) == 0 ||
                field.PrimaryKey || field.DataType == "AutoNumber")
            {
                EditTableDesign_Click(sender, e);
                return;
            }
        }
        catch (Exception ex)
        {
            ShowError("Could not inspect saved records", ex);
            return;
        }

        TableDesignerWindow designer;
        try { designer = new TableDesignerWindow(table, field) { Owner = this }; }
        catch (Exception ex) { ShowError("Unsupported column type", ex); return; }
        if (designer.ShowDialog() != true) return;
        var candidate = designer.Definition[0];
        if (AccessFieldMetadata.IsSameType(field, candidate))
        {
            MessageBox.Show(this,
                "The field already has that data type and size. " +
                "No database change or backup is necessary.",
                "No changes to save", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        Apply(() => _session.AlterFieldType(table, candidate),
            "Changed data type of " + field.Name + ".", table);
    }

    private void DropField_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table ||
            FieldsGrid.SelectedItem is not SchemaColumnInfo field)
        {
            MessageBox.Show(this, "Select a field row to delete.",
                "Select a field", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (field.PrimaryKey)
        {
            MessageBox.Show(this, "You cannot delete the primary key through this operation.",
                "Protected field", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Apply(() => _session.DropField(table, field.Name),
            "Deleted field " + field.Name + ".", table);
    }

    private void CreateIndex_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table || IndexColumnCombo.SelectedItem is not string column)
            return;
        string name = IndexNameBox.Text.Trim();
        bool unique = UniqueIndexCheck.IsChecked == true;
        Apply(() => _session.CreateIndex(table, name, column, unique),
            "Created index " + name + ".", table);
    }

    private void DropIndex_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table || IndexesCombo.SelectedItem is not string index)
            return;
        Apply(() => _session.DeleteIndex(table, index), "Removed index " + index + ".", table);
    }

    private void CreateRelationship_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string child ||
            ChildFieldCombo.SelectedItem is not string childField ||
            ParentTableCombo.SelectedItem is not string parent ||
            ParentFieldCombo.SelectedItem is not string parentField)
        {
            MessageBox.Show(this, "Select a source field, a referenced table, and its primary key field.",
                "Incomplete relationship", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        string name = RelationNameBox.Text.Trim();
        Apply(() => _session.AddRelationship(name, child, childField, parent, parentField),
            "Created relationship " + name + ".", child);
    }

    private void DeleteTable_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table) return;
        Apply(() => _session.DeleteTable(table), "Deleted table " + table + ".");
    }

    private void ShowError(string title, Exception ex)
    {
        MessageBox.Show(this, ex.Message, title, MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

