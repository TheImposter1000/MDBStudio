using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace MDBStudio;

/// <summary>Simple, explicit Access field types supported by the schema designer.</summary>
public enum AccessFieldKind
{
    AutoNumber,
    ShortText,
    LongText,
    Integer,
    SmallInteger,
    Byte,
    Double,
    Currency,
    DateTime,
    YesNo,
    Guid
}

public sealed record AccessFieldDefinition(
    string Name, AccessFieldKind Kind, int Length = 255,
    bool Required = false, bool PrimaryKey = false);

/// <summary>
/// Validates schema names and generates Access SQL without accepting arbitrary
/// SQL fragments from UI inputs. Data definitions only, never record values.
/// </summary>
public static class AccessSchemaSql
{
    public static string Name(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value != value.Trim() ||
            value.Length > 64 || value.Contains(']') || value.Contains('\0') ||
            value.Any(char.IsControl) || value.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) ||
            value.StartsWith("~", StringComparison.Ordinal))
            throw new ArgumentException(
                "Use a table, field, index or relationship name from 1 to 64 characters, " +
                "without control characters, brackets, or reserved MSys/~ prefixes.");
        return "[" + value + "]";
    }

    private static string TypeSql(AccessFieldDefinition field)
    {
        return field.Kind switch
        {
            AccessFieldKind.AutoNumber => "COUNTER",
            AccessFieldKind.ShortText => field.Length is >= 1 and <= 255
                ? "TEXT(" + field.Length + ")"
                : throw new ArgumentOutOfRangeException(nameof(field), "Short text fields must be 1 to 255 characters."),
            AccessFieldKind.LongText => "MEMO",
            AccessFieldKind.Integer => "LONG",
            AccessFieldKind.SmallInteger => "SHORT",
            AccessFieldKind.Byte => "BYTE",
            AccessFieldKind.Double => "DOUBLE",
            AccessFieldKind.Currency => "CURRENCY",
            AccessFieldKind.DateTime => "DATETIME",
            AccessFieldKind.YesNo => "YESNO",
            AccessFieldKind.Guid => "GUID",
            _ => throw new ArgumentException("Unsupported Access field type.")
        };
    }

    public static string Column(AccessFieldDefinition field)
    {
        string type = TypeSql(field);
        string name = Name(field.Name);
        if (field.Kind == AccessFieldKind.AutoNumber && field.Required == false)
            return name + " " + type;
        return name + " " + type + (field.Required && field.Kind != AccessFieldKind.AutoNumber
            ? " NOT NULL" : "");
    }

    public static string CreateTable(string table, IReadOnlyList<AccessFieldDefinition> fields)
    {
        string target = Name(table);
        if (fields.Count is < 1 or > 255)
            throw new InvalidOperationException("A table needs between 1 and 255 fields.");

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fields)
        {
            Name(field.Name);
            if (!seen.Add(field.Name))
                throw new InvalidOperationException("Duplicate field name: " + field.Name);
            _ = Column(field);
        }

        var autonumbers = fields.Where(f => f.Kind == AccessFieldKind.AutoNumber).ToArray();
        if (autonumbers.Length > 1)
            throw new InvalidOperationException("An Access table supports only one AutoNumber field.");
        if (autonumbers.Length == 1 && !autonumbers[0].PrimaryKey)
            throw new InvalidOperationException("Mark the AutoNumber field as the primary key.");

        var keys = fields.Where(f => f.PrimaryKey).ToArray();
        if (keys.Length == 0)
            throw new InvalidOperationException(
                "Choose at least one primary key field so MDB Studio can edit and save records in this table.");
        if (keys.Any(f => f.Kind == AccessFieldKind.LongText))
            throw new InvalidOperationException("Long text cannot be a primary key.");

        string primary = "CONSTRAINT " + Name("PK_" + TableConstraintSuffix(table)) +
            " PRIMARY KEY (" + string.Join(", ", keys.Select(f => Name(f.Name))) + ")";
        return "CREATE TABLE " + target + " (" +
            string.Join(", ", fields.Select(Column).Append(primary)) + ")";
    }

    private static string TableConstraintSuffix(string table)
    {
        // A constraint name must also stay within the 64-char Access limit.
        return table.Length > 61 ? table[..61] : table;
    }

    public static string AddField(string table, AccessFieldDefinition field)
    {
        if (field.PrimaryKey || field.Kind == AccessFieldKind.AutoNumber)
            throw new InvalidOperationException(
                "A new AutoNumber or primary key requires a new table design. " +
                "Add ordinary fields to existing tables; changing their keys is not supported here.");
        return "ALTER TABLE " + Name(table) + " ADD COLUMN " + Column(field);
    }

    public static string DropField(string table, string field) =>
        "ALTER TABLE " + Name(table) + " DROP COLUMN " + Name(field);

    public static string AlterFieldType(string table, AccessFieldDefinition field)
    {
        if (field.PrimaryKey || field.Kind == AccessFieldKind.AutoNumber)
            throw new InvalidOperationException(
                "Primary key and AutoNumber fields cannot be converted by this designer.");
        // ALTER COLUMN changes only the Access data type and optional text size.
        // Required/NULL properties must not be silently altered.
        return "ALTER TABLE " + Name(table) + " ALTER COLUMN " +
            Name(field.Name) + " " + TypeSql(field);
    }

    public static string CreateIndex(string table, string name,
        IReadOnlyList<string> columns, bool unique)
    {
        if (columns.Count == 0 || columns.Count > 10)
            throw new InvalidOperationException("Choose 1 to 10 index fields.");
        var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in columns)
        {
            Name(field);
            if (!distinct.Add(field))
                throw new InvalidOperationException("An index cannot repeat a field.");
        }
        return "CREATE " + (unique ? "UNIQUE " : "") + "INDEX " +
            Name(name) + " ON " + Name(table) + " (" +
            string.Join(", ", columns.Select(Name)) + ")";
    }

    public static string DropIndex(string table, string index) =>
        "DROP INDEX " + Name(index) + " ON " + Name(table);

    public static string AddRelationship(string constraint, string child,
        string childField, string parent, string parentField)
    {
        if (child.Equals(parent, StringComparison.OrdinalIgnoreCase) &&
            childField.Equals(parentField, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("A field cannot be its own relationship target.");
        return "ALTER TABLE " + Name(child) + " ADD CONSTRAINT " + Name(constraint) +
            " FOREIGN KEY (" + Name(childField) + ") REFERENCES " +
            Name(parent) + " (" + Name(parentField) + ")";
    }

    public static string DeleteTable(string table) => "DROP TABLE " + Name(table);
}
