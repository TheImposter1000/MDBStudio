using System.Data;

namespace MDBStudio;

/// <summary>
/// Read-only table filters compiled to parameterized ACE/OLE DB predicates.
/// Column names are validated against the actual MDB table schema.
/// </summary>
public enum TableFilterOperation
{
    Contains,
    Equals,
    StartsWith,
    EndsWith,
    NotEquals,
    IsBlank,
    IsNotBlank,
    GreaterThan,
    LessThan
}

public sealed record TableColumnFilter(string ColumnName, TableFilterOperation Operation, string Value);

internal sealed record FilterParameter(DataColumn Column, object Value);
internal sealed record CompiledTableFilters(string WhereClause, IReadOnlyList<FilterParameter> Parameters);

internal static class TableFilters
{
    public const int MaximumFilters = 12;

    public static bool CanFilter(DataColumn column) =>
        column.DataType != typeof(byte[]) &&
        (column.DataType == typeof(string) ||
         column.DataType == typeof(byte) ||
         column.DataType == typeof(short) ||
         column.DataType == typeof(int) ||
         column.DataType == typeof(long) ||
         column.DataType == typeof(float) ||
         column.DataType == typeof(double) ||
         column.DataType == typeof(decimal) ||
         column.DataType == typeof(bool) ||
         column.DataType == typeof(DateTime) ||
         column.DataType == typeof(Guid)) &&
        !column.ColumnName.Contains(']') && !column.ColumnName.Contains('\0');

    public static bool Supports(DataColumn column, TableFilterOperation operation)
    {
        if (!CanFilter(column)) return false;
        if (operation is TableFilterOperation.IsBlank or TableFilterOperation.IsNotBlank or
            TableFilterOperation.Equals or TableFilterOperation.NotEquals) return true;

        if (column.DataType == typeof(string))
            return operation is TableFilterOperation.Contains or
                TableFilterOperation.StartsWith or TableFilterOperation.EndsWith;

        return column.DataType != typeof(bool) &&
            column.DataType != typeof(Guid) &&
            operation is TableFilterOperation.GreaterThan or TableFilterOperation.LessThan;
    }

    public static CompiledTableFilters Compile(
        DataColumnCollection columns, IReadOnlyList<TableColumnFilter> filters)
    {
        if (filters.Count > MaximumFilters)
            throw new InvalidOperationException("Too many active filters. Remove a filter first.");

        var conditions = new List<string>();
        var values = new List<FilterParameter>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var filter in filters)
        {
            if (!columns.Contains(filter.ColumnName))
                throw new InvalidOperationException("The filter column no longer exists: " + filter.ColumnName);
            var col = columns[filter.ColumnName]!;
            if (!seen.Add(col.ColumnName))
                throw new InvalidOperationException("Only one active filter is permitted per column.");
            if (!Supports(col, filter.Operation))
                throw new InvalidOperationException(
                    "This filter condition is not supported for column " + col.ColumnName + ".");
            var field = "[" + col.ColumnName + "]";
            var operation = filter.Operation;
            if (operation == TableFilterOperation.IsBlank)
            {
                conditions.Add(col.DataType == typeof(string)
                    ? "(" + field + " IS NULL OR " + field + " = '')"
                    : field + " IS NULL");
                continue;
            }
            if (operation == TableFilterOperation.IsNotBlank)
            {
                conditions.Add(col.DataType == typeof(string)
                    ? "(" + field + " IS NOT NULL AND " + field + " <> '')"
                    : field + " IS NOT NULL");
                continue;
            }

            if (string.IsNullOrWhiteSpace(filter.Value))
                throw new FormatException("Enter a value or select the Is blank condition.");

            object value;
            try { value = SpreadsheetImportValue.ConvertToColumn(filter.Value, col); }
            catch (Exception ex) when (ex is FormatException or OverflowException)
            {
                throw new FormatException(
                    "Invalid filter value for " + col.ColumnName + ": " + ex.Message, ex);
            }

            switch (operation)
            {
                case TableFilterOperation.Contains:
                    conditions.Add("InStr(1, " + field + ", ?, 1) > 0");
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.StartsWith:
                    conditions.Add("InStr(1, " + field + ", ?, 1) = 1");
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.EndsWith:
                    conditions.Add("Right(" + field + ", Len(?)) = ?");
                    values.Add(new FilterParameter(col, value));
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.Equals:
                    conditions.Add(field + " = ?");
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.NotEquals:
                    conditions.Add(field + " <> ?");
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.GreaterThan:
                    conditions.Add(field + " > ?");
                    values.Add(new FilterParameter(col, value));
                    break;
                case TableFilterOperation.LessThan:
                    conditions.Add(field + " < ?");
                    values.Add(new FilterParameter(col, value));
                    break;
            }
        }

        return new CompiledTableFilters(
            conditions.Count == 0 ? "" : " WHERE " + string.Join(" AND ", conditions.Select(c => "(" + c + ")")),
            values);
    }
}
