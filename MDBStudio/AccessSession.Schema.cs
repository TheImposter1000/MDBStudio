using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace MDBStudio;

public sealed record SchemaColumnInfo(
    string Name, string DataType, int Length, bool Required, bool PrimaryKey);

public sealed record SchemaTableInfo(
    string Name, IReadOnlyList<SchemaColumnInfo> Columns, IReadOnlyList<string> IndexNames);

/// <summary>Access table-design operations; every DDL statement is preceded by a new backup.</summary>
public sealed partial class AccessSession
{
    public SchemaTableInfo InspectTable(string table)
    {
        AccessSchemaSql.Name(table);
        if (!ListTables().Contains(table, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The selected table does not exist.");
        using var connection = Open();
        DataTable schema = ReadSchema(connection, table);
        var keys = FindPrimaryKeys(connection, table).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var columns = schema.Columns.Cast<DataColumn>()
            .Select(c => new SchemaColumnInfo(
                c.ColumnName, FriendlyType(c), c.MaxLength, !c.AllowDBNull,
                keys.Contains(c.ColumnName))).ToArray();
        return new SchemaTableInfo(table, columns, GetIndexes(connection, table));
    }

    private static string FriendlyType(DataColumn column)
    {
        Type type = column.DataType;
        if (column.AutoIncrement) return "AutoNumber";
        if (type == typeof(string)) return column.MaxLength is > 0 and <= 255 ? "Short Text" : "Long Text";
        if (type == typeof(int)) return "Integer (32-bit)";
        if (type == typeof(short)) return "Small Integer";
        if (type == typeof(byte)) return "Byte";
        if (type == typeof(bool)) return "Yes/No";
        if (type == typeof(DateTime)) return "Date/Time";
        if (type == typeof(double)) return "Double";
        if (type == typeof(float)) return "Single";
        if (type == typeof(decimal)) return "Decimal/Currency";
        if (type == typeof(Guid)) return "GUID";
        if (type == typeof(byte[])) return "Binary/OLE";
        return type.Name;
    }

    private static IReadOnlyList<string> GetIndexes(OleDbConnection connection, string table)
    {
        try
        {
            var data = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Indexes, null);
            if (data is null || !data.Columns.Contains("INDEX_NAME") ||
                !data.Columns.Contains("TABLE_NAME")) return Array.Empty<string>();
            return data.Rows.Cast<DataRow>()
                .Where(r => string.Equals(Convert.ToString(r["TABLE_NAME"], CultureInfo.InvariantCulture),
                    table, StringComparison.OrdinalIgnoreCase))
                .Where(r => !data.Columns.Contains("PRIMARY_KEY") ||
                    r["PRIMARY_KEY"] == DBNull.Value || !Convert.ToBoolean(r["PRIMARY_KEY"], CultureInfo.InvariantCulture))
                .Select(r => Convert.ToString(r["INDEX_NAME"], CultureInfo.InvariantCulture) ?? "")
                .Where(n => n.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        catch (OleDbException)
        {
            return Array.Empty<string>();
        }
    }

    private void ApplySchemaSql(string sql)
    {
        // DDL in Access may auto-commit and cannot reliably be rolled back.
        // Make a fresh restorable file copy before every structure change.
        Backup();
        using var connection = Open();
        using var command = new OleDbCommand(sql, connection);
        command.ExecuteNonQuery();
    }

    public void CreateTable( string table, IReadOnlyList<AccessFieldDefinition> fields)
    {
        string sql = AccessSchemaSql.CreateTable(table, fields);
        if (ListTables().Contains(table, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("A table with that name already exists.");
        ApplySchemaSql(sql);
    }

    public void AddField(string table, AccessFieldDefinition field)
    {
        string sql = AccessSchemaSql.AddField(table, field);
        var structure = InspectTable(table);
        if (structure.Columns.Any(c => c.Name.Equals(field.Name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("That field already exists in the table.");
        ApplySchemaSql(sql);
    }

    public void CreateIndex(string table, string indexName, string fieldName, bool unique)
    {
        var structure = InspectTable(table);
        if (!structure.Columns.Any(c => c.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("Choose a field that exists in this table.");
        if (structure.IndexNames.Contains(indexName, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("This table already has an index with that name.");
        string sql = AccessSchemaSql.CreateIndex(table, indexName, new[] { fieldName }, unique);
        ApplySchemaSql(sql);
    }

    public void DeleteIndex(string table, string indexName)
    {
        var structure = InspectTable(table);
        if (!structure.IndexNames.Contains(indexName, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The index is missing or is a protected primary key index.");
        ApplySchemaSql(AccessSchemaSql.DropIndex(table, indexName));
    }

    public void AddRelationship(string constraint, string childTable, string childField,
        string parentTable, string parentField)
    {
        string sql = AccessSchemaSql.AddRelationship(
            constraint, childTable, childField, parentTable, parentField);
        var child = InspectTable(childTable);
        var parent = InspectTable(parentTable);
        var childColumn = child.Columns.FirstOrDefault(c =>
            c.Name.Equals(childField, StringComparison.OrdinalIgnoreCase));
        var parentColumn = parent.Columns.FirstOrDefault(c =>
            c.Name.Equals(parentField, StringComparison.OrdinalIgnoreCase));
        if (childColumn is null || parentColumn is null)
            throw new InvalidOperationException("Both relationship fields must exist.");
        if (!string.Equals(childColumn.DataType, parentColumn.DataType, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "Relationship fields must have matching data types. " +
                "The referenced field must also be a primary key or unique index.");
        ApplySchemaSql(sql);
    }

    public void DeleteTable(string table)
    {
        string sql = AccessSchemaSql.DeleteTable(table);
        if (!ListTables().Contains(table, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The table no longer exists.");
        ApplySchemaSql(sql);
    }
}
