using System.Windows;
using System.Windows.Controls;
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
            FieldsGrid.ItemsSource = info.Columns;
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
        var designer = new TableDesignerWindow { Owner = this };
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
        if (MessageBox.Show(this, "Delete index '" + index + "' from '" + table +
            "'? This may affect queries or relationships. A backup is created first.",
            "Confirm index removal", MessageBoxButton.YesNo, MessageBoxImage.Warning,
            MessageBoxResult.No) != MessageBoxResult.Yes) return;
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
        if (MessageBox.Show(this,
            "Create foreign key '" + name + "' from " + child + "." + childField +
            " to " + parent + "." + parentField +
            "?\n\nExisting data must satisfy the relationship. A backup will be created first.",
            "Confirm relationship", MessageBoxButton.YesNo, MessageBoxImage.Question,
            MessageBoxResult.No) != MessageBoxResult.Yes) return;

        Apply(() => _session.AddRelationship(name, child, childField, parent, parentField),
            "Created relationship " + name + ".", child);
    }

    private void DeleteTable_Click(object sender, RoutedEventArgs e)
    {
        if (SelectedTable is not string table) return;
        var prompt = new ExactTableConfirmation(table) { Owner = this };
        if (prompt.ShowDialog() != true) return;
        Apply(() => _session.DeleteTable(table), "Deleted table " + table + ".");
    }

    private void ShowError(string title, Exception ex)
    {
        MessageBox.Show(this, ex.Message, title, MessageBoxButton.OK, MessageBoxImage.Error);
    }
}

/// <summary>Explicit exact-name confirmation for irreversible Access DROP TABLE.</summary>
internal sealed class ExactTableConfirmation : Window
{
    private readonly TextBox _typed = new();
    private readonly string _expected;

    public ExactTableConfirmation(string table)
    {
        _expected = table;
        Title = "Confirm table deletion";
        Width = 485;
        Height = 245;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;
        Background = Brushes.White;
        var content = new StackPanel { Margin = new Thickness(20) };
        content.Children.Add(new TextBlock
        {
            Text = "Deleting a table removes all its records and may break relationships. " +
                   "MDB Studio creates a separate file backup before deletion.",
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 14)
        });
        content.Children.Add(new TextBlock
        {
            Text = "Type the exact table name to confirm: " + table,
            FontWeight = FontWeights.SemiBold,
            TextWrapping = TextWrapping.Wrap
        });
        _typed.Margin = new Thickness(0, 8, 0, 14);
        content.Children.Add(_typed);
        var row = new StackPanel
        {
            Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right
        };
        var yes = new Button
        {
            Content = "Delete table", Padding = new Thickness(13, 8, 13, 8),
            Margin = new Thickness(0, 0, 8, 0)
        };
        yes.Click += (_, _) =>
        {
            if (_typed.Text != _expected)
                MessageBox.Show(this, "Table name does not match. Nothing was deleted.",
                    "Confirmation required", MessageBoxButton.OK, MessageBoxImage.Warning);
            else DialogResult = true;
        };
        row.Children.Add(yes);
        row.Children.Add(new Button
            { Content = "Cancel", IsCancel = true, Padding = new Thickness(13, 8, 13, 8) });
        content.Children.Add(row);
        Content = content;
    }
}
