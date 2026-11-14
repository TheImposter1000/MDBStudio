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
        DataTable? columnsSchema = null;
        try
        {
            columnsSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns,
                new object?[] { null, null, table, null });
        }
        catch (OleDbException) { /* Fall back to FillSchema if the provider cannot supply metadata. */ }

        var keys = FindPrimaryKeys(connection, table).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var columns = schema.Columns.Cast<DataColumn>().Select(c =>
        {
            var metadata = AccessFieldMetadata.ReadColumnMetadata(columnsSchema, c.ColumnName);
            string type = AccessFieldMetadata.FriendlyType(c, metadata.Type, metadata.Length);
            int length = type == "Short Text"
                ? AccessFieldMetadata.TextLength(c, metadata.Length) : c.MaxLength;
            return new SchemaColumnInfo(c.ColumnName, type, length,
                !c.AllowDBNull, keys.Contains(c.ColumnName));
        }).ToArray();
        return new SchemaTableInfo(table, columns, GetIndexes(connection, table));
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
        static string RelatedType(string type) =>
            type.Equals("AutoNumber", StringComparison.OrdinalIgnoreCase)
                ? "Integer (32-bit)" : type;
        if (!string.Equals(RelatedType(childColumn.DataType), RelatedType(parentColumn.DataType),
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "Relationship fields must have matching data types (32-bit Integer is compatible " +
                "with AutoNumber). The referenced field must also be a primary key or unique index.");
        ApplySchemaSql(sql);
    }

    public IReadOnlyDictionary<string, Type> QueryColumnTypes(string table)
    {
        AccessSchemaSql.Name(table);
        if (!ListTables().Contains(table, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The query table does not exist.");
        using var conn = Open();
        return ReadSchema(conn, table).Columns.Cast<DataColumn>()
            .ToDictionary(c => c.ColumnName, c => c.DataType,
                StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<string> ListSavedQueries()
    {
        using var conn = Open();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var views = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Views, null);
            if (views is not null && views.Columns.Contains("TABLE_NAME"))
                foreach (DataRow row in views.Rows)
                {
                    string name = Convert.ToString(row["TABLE_NAME"], CultureInfo.InvariantCulture) ?? "";
                    if (name.Length != 0 && !name.StartsWith("~", StringComparison.Ordinal))
                        names.Add(name);
                }
        }
        catch (OleDbException) { }
        try
        {
            var procedures = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Procedures, null);
            if (procedures is not null && procedures.Columns.Contains("PROCEDURE_NAME"))
                foreach (DataRow row in procedures.Rows)
                {
                    string name = Convert.ToString(row["PROCEDURE_NAME"], CultureInfo.InvariantCulture) ?? "";
                    if (name.EndsWith(";1", StringComparison.Ordinal)) name = name[..^2];
                    if (name.Length != 0 && !name.StartsWith("~", StringComparison.Ordinal))
                        names.Add(name);
                }
        }
        catch (OleDbException) { }
        return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public void SaveSelectView(string name, string sourceTable,
        IReadOnlyList<VisualQueryField> fields, VisualQueryJoin? join, bool distinct)
    {
        string sql = VisualQuerySql.Build(sourceTable, fields, join, distinct);
        string id = AccessSchemaSql.Name(name);
        if (ListTables().Contains(name, StringComparer.OrdinalIgnoreCase) ||
            ListSavedQueries().Contains(name, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("A table or query with that name already exists.");

        // CREATE VIEW requires Microsoft ACE/Jet; backups are required for all schema writes.
        ApplySchemaSql("CREATE VIEW " + id + " AS " + sql);
    }

    public void DropField(string table, string column)
    {
        var structure = InspectTable(table);
        var field = structure.Columns.FirstOrDefault(f =>
            f.Name.Equals(column, StringComparison.OrdinalIgnoreCase)) ??
            throw new InvalidOperationException("The selected field does not exist.");
        if (field.PrimaryKey)
            throw new InvalidOperationException("A primary key cannot be removed. " +
                "Change the table's key design before removing this column.");
        if (structure.Columns.Count <= 1)
            throw new InvalidOperationException("A table must retain at least one field.");
        // DDL will be blocked by ACE for referenced/indexed fields.
        ApplySchemaSql(AccessSchemaSql.DropField(table, field.Name));
    }

    public void AlterFieldType(string table, AccessFieldDefinition target)
    {
        var structure = InspectTable(table);
        var existing = structure.Columns.FirstOrDefault(f =>
            f.Name.Equals(target.Name, StringComparison.OrdinalIgnoreCase)) ??
            throw new InvalidOperationException("The selected field does not exist.");
        if (existing.PrimaryKey || existing.DataType == "AutoNumber")
            throw new InvalidOperationException("Changing the type of a primary key or AutoNumber field is not supported.");
        if (target.Kind == AccessFieldKind.AutoNumber || target.PrimaryKey)
            throw new InvalidOperationException("Changing a field to an AutoNumber or primary key is not supported.");

        // This operation can truncate or fail if existing values are incompatible.
        // Caller must require explicit confirmation. Backup always precedes DDL.
        ApplySchemaSql(AccessSchemaSql.AlterFieldType(table, target));
    }

    public void DeleteTable(string table)
    {
        string sql = AccessSchemaSql.DeleteTable(table);
        if (!ListTables().Contains(table, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The table no longer exists.");
        ApplySchemaSql(sql);
    }
}
