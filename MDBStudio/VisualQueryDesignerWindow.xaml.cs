using System.Collections.ObjectModel;
using System.Data;
using System.Windows;
using System.Windows.Controls;

namespace MDBStudio;

public sealed class EditableVisualQueryField
{
    public string Table { get; init; } = "";
    public string Column { get; init; } = "";
    public Type DataType { get; init; } = typeof(string);
    public bool Show { get; set; } = true;
    public VisualQueryCriteria Criteria { get; set; }
    public string Value { get; set; } = "";
    public VisualQuerySort Sort { get; set; }

    public VisualQueryField ToField() => new(
        Table, Column, DataType, Show, Criteria, Value, Sort);
}

public partial class VisualQueryDesignerWindow : Window
{
    private sealed record FieldOption(string Table, string Column, Type DataType)
    {
        public string Label => Table + "." + Column;
    }

    private readonly AccessSession _session;
    private readonly ObservableCollection<EditableVisualQueryField> _fields = new();
    private bool _ready;
    private const string NoJoin = "(No second table)";

    public string? ResultSql { get; private set; }

    public VisualQueryDesignerWindow(AccessSession session, string? preferredTable = null)
    {
        InitializeComponent();
        _session = session;
        DesignGrid.ItemsSource = _fields;
        CriteriaColumn.ItemsSource = Enum.GetValues<VisualQueryCriteria>();
        SortColumn.ItemsSource = Enum.GetValues<VisualQuerySort>();
        JoinModeCombo.ItemsSource = new[] { "Inner join", "Left join" };
        JoinModeCombo.SelectedIndex = 0;
        SourceTableCombo.ItemsSource = session.ListTables();
        _ready = true;

        if (preferredTable is not null &&
            SourceTableCombo.Items.Contains(preferredTable))
            SourceTableCombo.SelectedItem = preferredTable;
        else if (SourceTableCombo.Items.Count > 0)
            SourceTableCombo.SelectedIndex = 0;

        ReloadSavedQueries();
        SaveNameBox.Text = "qryNewQuery";
    }

    private void ReloadSavedQueries()
    {
        try
        {
            SavedQueriesCombo.ItemsSource = _session.ListSavedQueries();
        }
        catch (Exception ex)
        {
            StatusLabel.Text = "Saved queries unavailable: " + ex.Message;
        }
    }

    private string? SelectedSource => SourceTableCombo.SelectedItem as string;
    private string? SelectedJoin => JoinTableCombo.SelectedItem as string is { } s && s != NoJoin
        ? s : null;

    private void SourceTableChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_ready) return;
        string? source = SelectedSource;
        _fields.Clear();
        try
        {
            var choices = new List<string> { NoJoin };
            if (source is not null)
                choices.AddRange(_session.ListTables().Where(t =>
                    !t.Equals(source, StringComparison.OrdinalIgnoreCase)));
            JoinTableCombo.ItemsSource = choices;
            JoinTableCombo.SelectedIndex = 0;
            UpdateAvailableFields();
        }
        catch (Exception ex) { Error("Could not list query fields", ex); }
    }

    private void JoinTableChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_ready) return;
        string? join = SelectedJoin;
        _fields.ToList().ForEach(f =>
        {
            if (!f.Table.Equals(SelectedSource, StringComparison.OrdinalIgnoreCase) &&
                !f.Table.Equals(join, StringComparison.OrdinalIgnoreCase))
                _fields.Remove(f);
        });
        UpdateAvailableFields();
    }

    private void UpdateAvailableFields()
    {
        if (!_ready || SelectedSource is not string source) return;
        try
        {
            var baseFields = _session.QueryColumnTypes(source);
            SourceKeyCombo.ItemsSource = baseFields.Keys.ToArray();
            SourceKeyCombo.SelectedIndex = baseFields.Count == 0 ? -1 : 0;
            var available = baseFields.Select(k => new FieldOption(source, k.Key, k.Value)).ToList();
            if (SelectedJoin is string other)
            {
                var joined = _session.QueryColumnTypes(other);
                JoinKeyCombo.ItemsSource = joined.Keys.ToArray();
                JoinKeyCombo.SelectedIndex = joined.Count == 0 ? -1 : 0;
                available.AddRange(joined.Select(k => new FieldOption(other, k.Key, k.Value)));
            }
            else
            {
                JoinKeyCombo.ItemsSource = null;
            }
            JoinKeyCombo.IsEnabled = SelectedJoin is not null;
            SourceKeyCombo.IsEnabled = SelectedJoin is not null;
            JoinModeCombo.IsEnabled = SelectedJoin is not null;
            AvailableFieldCombo.ItemsSource = available;
            if (available.Count > 0) AvailableFieldCombo.SelectedIndex = 0;
        }
        catch (Exception ex) { Error("Could not read query table fields", ex); }
    }

    private void AddField_Click(object sender, RoutedEventArgs e)
    {
        if (AvailableFieldCombo.SelectedItem is not FieldOption option) return;
        if (_fields.Any(f => f.Table.Equals(option.Table, StringComparison.OrdinalIgnoreCase) &&
            f.Column.Equals(option.Column, StringComparison.OrdinalIgnoreCase)))
        {
            StatusLabel.Text = "This field is already in the query. Set its Show, Criteria, or Sort values.";
            return;
        }
        var field = new EditableVisualQueryField
        {
            Table = option.Table, Column = option.Column, DataType = option.DataType
        };
        _fields.Add(field);
        DesignGrid.SelectedItem = field;
        DesignGrid.ScrollIntoView(field);
    }

    private void RemoveField_Click(object sender, RoutedEventArgs e)
    {
        if (DesignGrid.SelectedItem is EditableVisualQueryField field)
            _fields.Remove(field);
    }

    private (string Source, VisualQueryField[] Fields, VisualQueryJoin? Join, bool Distinct)
        ReadDesign()
    {
        DesignGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        DesignGrid.CommitEdit(DataGridEditingUnit.Row, true);
        string source = SelectedSource ??
            throw new InvalidOperationException("Choose a source table.");
        VisualQueryJoin? join = null;
        if (SelectedJoin is string secondary)
        {
            string primaryField = SourceKeyCombo.SelectedItem as string ??
                throw new InvalidOperationException("Choose the source join field.");
            string secondaryField = JoinKeyCombo.SelectedItem as string ??
                throw new InvalidOperationException("Choose the joined table field.");
            Type primaryType = _session.QueryColumnTypes(source)[primaryField];
            Type secondaryType = _session.QueryColumnTypes(secondary)[secondaryField];
            if (primaryType != secondaryType)
                throw new InvalidOperationException(
                    "Join fields must have compatible data types. " +
                    "For example, two 32-bit integer fields can be joined.");
            join = new VisualQueryJoin(source, primaryField, secondary, secondaryField,
                JoinModeCombo.SelectedIndex == 1);
        }
        return (source, _fields.Select(f => f.ToField()).ToArray(),
            join, DistinctCheck.IsChecked == true);
    }

    private string Generate()
    {
        var design = ReadDesign();
        return VisualQuerySql.Build(design.Source, design.Fields, design.Join, design.Distinct);
    }

    private void Preview_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SqlPreviewBox.Text = Generate();
            StatusLabel.Text = "Read-only SELECT ready. Run or save it.";
        }
        catch (Exception ex) { Error("Invalid query design", ex); }
    }

    private void Run_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ResultSql = Generate();
            DialogResult = true;
        }
        catch (Exception ex) { Error("Invalid query design", ex); }
    }

    private void SaveQuery_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var design = ReadDesign();
            string sql = VisualQuerySql.Build(design.Source, design.Fields,
                design.Join, design.Distinct);
            string name = SaveNameBox.Text.Trim();
            _session.SaveSelectView(name, design.Source, design.Fields,
                design.Join, design.Distinct);
            SqlPreviewBox.Text = sql;
            ReloadSavedQueries();
            SavedQueriesCombo.SelectedItem = name;
            StatusLabel.Text = "Saved " + name + " as a query in the MDB. A backup was created.";
        }
        catch (Exception ex) { Error("Could not save Access query", ex); }
    }

    private void OpenSavedQuery_Click(object sender, RoutedEventArgs e)
    {
        if (SavedQueriesCombo.SelectedItem is not string query)
        {
            StatusLabel.Text = "Select a saved query to open.";
            return;
        }
        try
        {
            ResultSql = "SELECT * FROM " + AccessSchemaSql.Name(query);
            DialogResult = true;
        }
        catch (Exception ex) { Error("Invalid saved query", ex); }
    }

    private void Error(string caption, Exception ex)
    {
        MessageBox.Show(this, ex.Message, caption, MessageBoxButton.OK, MessageBoxImage.Warning);
    }
}
