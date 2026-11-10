using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace MDBStudio;

public sealed class EditableSchemaField
{
    public string Name { get; set; } = "";
    public AccessFieldKind Kind { get; set; } = AccessFieldKind.ShortText;
    public int Length { get; set; } = 255;
    public bool Required { get; set; }
    public bool PrimaryKey { get; set; }
    public AccessFieldDefinition ToDefinition() => new(Name.Trim(), Kind, Length, Required, PrimaryKey);
}

public partial class TableDesignerWindow : Window
{
    private readonly bool _addingColumn;
    private readonly bool _changingType;
    private readonly ObservableCollection<EditableSchemaField> _fields = new();

    public string TableName => TableNameBox.Text.Trim();
    public IReadOnlyList<AccessFieldDefinition> Definition { get; private set; } =
        Array.Empty<AccessFieldDefinition>();

    public TableDesignerWindow(string? existingTable = null, SchemaColumnInfo? existingField = null)
    {
        InitializeComponent();
        _addingColumn = existingTable is not null && existingField is null;
        _changingType = existingField is not null;
        FieldTypeColumn.ItemsSource = Enum.GetValues<AccessFieldKind>();
        FieldsGrid.ItemsSource = _fields;
        if (_changingType)
        {
            Heading.Text = "Change field type in " + existingTable;
            HelpText.Text = "Changing a type or reducing text length may truncate data, " +
                "invalidate queries or break relationships. MDB Studio backs up the database before applying it.";
            TableNameBox.Text = existingTable!;
            TableNameBox.IsReadOnly = true;
            AddFieldButton.IsEnabled = false;
            RemoveFieldButton.IsEnabled = false;
            FieldNameColumn.IsReadOnly = true;
            RequiredColumn.IsReadOnly = true;
            PrimaryKeyColumn.IsReadOnly = true;
            SaveButton.Content = "Review type change";
            _fields.Add(new EditableSchemaField
            {
                Name = existingField!.Name,
                Kind = KindFromSchema(existingField.DataType),
                Length = existingField.Length is > 0 and <= 255 ? existingField.Length : 255,
                Required = existingField.Required,
                PrimaryKey = false
            });
        }
        else if (_addingColumn)
        {
            Heading.Text = "Add a field to " + existingTable;
            HelpText.Text =
                "Choose a name, type and optional required setting. " +
                "Existing records must accept the new field. Fields cannot be renamed or converted here.";
            TableNameBox.Text = existingTable!;
            TableNameBox.IsReadOnly = true;
            AddFieldButton.IsEnabled = false;
            RemoveFieldButton.IsEnabled = false;
            SaveButton.Content = "Add field";
            _fields.Add(new EditableSchemaField { Name = "NewField" });
        }
        else
        {
            TableNameBox.Text = "NewTable";
            _fields.Add(new EditableSchemaField
                { Name = "ID", Kind = AccessFieldKind.AutoNumber, PrimaryKey = true });
            _fields.Add(new EditableSchemaField
                { Name = "Name", Kind = AccessFieldKind.ShortText, Length = 100 });
        }
    }

    private static AccessFieldKind KindFromSchema(string type) => type switch
    {
        "Short Text" => AccessFieldKind.ShortText,
        "Long Text" => AccessFieldKind.LongText,
        "Integer (32-bit)" => AccessFieldKind.Integer,
        "Small Integer" => AccessFieldKind.SmallInteger,
        "Byte" => AccessFieldKind.Byte,
        "Double" => AccessFieldKind.Double,
        "Yes/No" => AccessFieldKind.YesNo,
        "Date/Time" => AccessFieldKind.DateTime,
        "GUID" => AccessFieldKind.Guid,
        "Decimal/Currency" => AccessFieldKind.Currency,
        _ => throw new InvalidOperationException(
            "This existing column type cannot be changed by the visual designer.")
    };

    private void AddField_Click(object sender, RoutedEventArgs e)
    {
        var field = new EditableSchemaField();
        _fields.Add(field);
        FieldsGrid.SelectedItem = field;
        FieldsGrid.ScrollIntoView(field);
    }

    private void RemoveField_Click(object sender, RoutedEventArgs e)
    {
        if (FieldsGrid.SelectedItem is EditableSchemaField field && !_addingColumn)
            _fields.Remove(field);
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            FieldsGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            FieldsGrid.CommitEdit(DataGridEditingUnit.Row, true);
            _ = AccessSchemaSql.Name(TableName);
            var definition = _fields.Select(f => f.ToDefinition()).ToArray();
            if (_changingType)
            {
                if (definition.Length != 1)
                    throw new InvalidOperationException("Choose one field to change.");
                _ = AccessSchemaSql.AlterFieldType(TableName, definition[0]);
            }
            else if (_addingColumn)
            {
                if (definition.Length != 1)
                    throw new InvalidOperationException("Select exactly one field to add.");
                _ = AccessSchemaSql.AddField(TableName, definition[0]);
            }
            else
                _ = AccessSchemaSql.CreateTable(TableName, definition);
            Definition = definition;
            DialogResult = true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "Table design validation",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
