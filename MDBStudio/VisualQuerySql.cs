using System.Globalization;

namespace MDBStudio;

public enum VisualQueryCriteria
{
    None,
    Equals,
    NotEquals,
    Contains,
    StartsWith,
    GreaterThan,
    LessThan,
    IsNull,
    IsNotNull
}

public enum VisualQuerySort
{
    None,
    Ascending,
    Descending
}

public sealed record VisualQueryField(
    string Table, string Column, Type DataType, bool Show,
    VisualQueryCriteria Criteria = VisualQueryCriteria.None,
    string Value = "", VisualQuerySort Sort = VisualQuerySort.None);

public sealed record VisualQueryJoin(
    string BaseTable, string BaseField, string JoinedTable,
    string JoinedField, bool LeftJoin);

/// <summary>
/// Creates Access SELECT statements from individually validated table/field choices.
/// Field references are constrained to those retrieved from the open MDB schema.
/// </summary>
public static class VisualQuerySql
{
    public static string Build(string mainTable, IReadOnlyList<VisualQueryField> fields,
        VisualQueryJoin? join = null, bool distinct = false)
    {
        AccessSchemaSql.Name(mainTable);
        if (fields.Count == 0 || fields.Count > 255)
            throw new InvalidOperationException("Select 1 to 255 fields.");

        string from = " FROM " + AccessSchemaSql.Name(mainTable);
        if (join is not null)
        {
            if (!join.BaseTable.Equals(mainTable, StringComparison.OrdinalIgnoreCase) ||
                join.JoinedTable.Equals(mainTable, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The join must link two different tables.");
            from += (join.LeftJoin ? " LEFT JOIN " : " INNER JOIN ") +
                AccessSchemaSql.Name(join.JoinedTable) + " ON " +
                Reference(join.BaseTable, join.BaseField) + " = " +
                Reference(join.JoinedTable, join.JoinedField);
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in fields)
        {
            if (!field.Table.Equals(mainTable, StringComparison.OrdinalIgnoreCase) &&
                (join is null || !field.Table.Equals(join.JoinedTable,
                    StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("A selected field is from a table outside the query.");
            string key = field.Table + "\0" + field.Column;
            if (!seen.Add(key))
                throw new InvalidOperationException("A field cannot occur twice in the query design.");
            _ = Reference(field.Table, field.Column);
        }

        var visible = fields.Where(f => f.Show).ToArray();
        if (visible.Length == 0)
            throw new InvalidOperationException("Select at least one visible field.");
        // A JOIN may have identically named fields: use explicit unique aliases.
        string[] projections = visible.Select(f => join is null
            ? Reference(f.Table, f.Column)
            : Reference(f.Table, f.Column) + " AS " +
                AccessSchemaSql.Name(Alias(f.Table, f.Column))).ToArray();
        if (join is not null && visible.Select(f => Alias(f.Table, f.Column))
            .Distinct(StringComparer.OrdinalIgnoreCase).Count() != visible.Length)
            throw new InvalidOperationException(
                "The chosen table and field names produce ambiguous output aliases.");

        var conditions = fields.Where(f => f.Criteria != VisualQueryCriteria.None)
            .Select(Condition).ToArray();
        var sort = fields.Where(f => f.Sort != VisualQuerySort.None)
            .Select(f => Reference(f.Table, f.Column) +
                (f.Sort == VisualQuerySort.Descending ? " DESC" : " ASC")).ToArray();

        return "SELECT " + (distinct ? "DISTINCT " : "") +
            string.Join(", ", projections) + from +
            (conditions.Length == 0 ? "" : " WHERE " + string.Join(" AND ", conditions)) +
            (sort.Length == 0 ? "" : " ORDER BY " + string.Join(", ", sort));
    }

    public static string Reference(string table, string field) =>
        AccessSchemaSql.Name(table) + "." + AccessSchemaSql.Name(field);

    private static string Alias(string table, string field)
    {
        // Use deterministic unique, readable aliases for joined result columns.
        string raw = table + "_" + field;
        if (raw.Length <= 64) return raw;
        // At most 64 characters in Access names; reject collisions in Build().
        return raw[..64];
    }

    private static string Condition(VisualQueryField field)
    {
        string reference = Reference(field.Table, field.Column);
        if (field.Criteria == VisualQueryCriteria.IsNull) return reference + " IS NULL";
        if (field.Criteria == VisualQueryCriteria.IsNotNull) return reference + " IS NOT NULL";
        if (string.IsNullOrWhiteSpace(field.Value))
            throw new InvalidOperationException(
                "Enter a criterion for " + field.Table + "." + field.Column +
                ", or choose IsNull/IsNotNull.");

        if (field.Criteria is VisualQueryCriteria.Contains or VisualQueryCriteria.StartsWith &&
            field.DataType != typeof(string))
            throw new InvalidOperationException("Text search requires a text field.");

        string literal = ValueLiteral(field.Value, field.DataType);
        return field.Criteria switch
        {
            VisualQueryCriteria.Equals => reference + " = " + literal,
            VisualQueryCriteria.NotEquals => reference + " <> " + literal,
            VisualQueryCriteria.Contains => "InStr(1, " + reference + ", " + literal + ", 1) > 0",
            VisualQueryCriteria.StartsWith => "InStr(1, " + reference + ", " + literal + ", 1) = 1",
            VisualQueryCriteria.GreaterThan => Compare(reference, literal, field.DataType, ">"),
            VisualQueryCriteria.LessThan => Compare(reference, literal, field.DataType, "<"),
            _ => throw new InvalidOperationException("Unsupported criterion.")
        };
    }

    private static string Compare(string reference, string literal, Type type, string op)
    {
        if (type == typeof(bool) || type == typeof(Guid) || type == typeof(byte[]))
            throw new InvalidOperationException("This data type cannot be ordered in a query criterion.");
        return reference + " " + op + " " + literal;
    }

    private static string ValueLiteral(string value, Type type)
    {
        if (type == typeof(string))
            return "'" + value.Replace("'", "''") + "'";
        if (type == typeof(bool))
        {
            if (value.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                value == "-1" || value == "1") return "True";
            if (value.Equals("false", StringComparison.OrdinalIgnoreCase) ||
                value.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                value == "0") return "False";
            throw new FormatException("Enter True, False, Yes, No, 1, 0 or -1.");
        }
        if (type == typeof(DateTime))
        {
            if (!DateTime.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out DateTime date))
                throw new FormatException("Enter a valid date/time.");
            return "#" + date.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture) + "#";
        }
        if (type == typeof(Guid))
        {
            if (!Guid.TryParse(value, out Guid guid))
                throw new FormatException("Enter a valid GUID.");
            return "'" + guid.ToString("D") + "'";
        }
        if (type == typeof(byte[]))
            throw new InvalidOperationException("Binary fields cannot be compared in the query designer.");

        object number;
        try
        {
            number = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException or InvalidCastException)
        {
            throw new FormatException("Enter a valid numeric value for this column.", ex);
        }
        return Convert.ToString(number, CultureInfo.InvariantCulture) ??
            throw new FormatException("Invalid numeric value.");
    }
}
