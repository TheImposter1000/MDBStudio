using System.Data;

namespace MDBStudio;

/// <summary>
/// Stages deletion of existing database records without committing to Access.
/// A row can be targeted by its row header or by any selected cell.
/// </summary>
public static class RecordDeletion
{
    public static IReadOnlyList<DataRow> EligibleRows(IEnumerable<DataRow> selected)
    {
        // A multiple-cell selection commonly references the same record more
        // than once. Never count or delete a record more than once.
        return selected.Where(row => row.RowState != DataRowState.Detached &&
                                     row.RowState != DataRowState.Deleted)
            .Distinct()
            .ToArray();
    }

    /// <summary>
    /// Generates a value-based predicate for tables without a primary key.
    /// Long text, binary, and unsupported columns cannot safely be compared
    /// by Access SQL and are excluded. Before deleting, the caller MUST verify
    /// that the predicate matches exactly the selected number of saved rows.
    /// </summary>
    internal sealed record KeylessMatch(string Where, IReadOnlyList<DataColumn> Parameters, IReadOnlyList<DataColumn> ComparedColumns);

    internal static KeylessMatch BuildKeylessMatch(DataRow row)
    {
        if (!row.HasVersion(DataRowVersion.Original))
            throw new InvalidOperationException("Only saved records can be matched for deletion.");

        var columns = row.Table.Columns.Cast<DataColumn>()
            .Where(c => !c.ColumnName.Contains(']') && !c.ColumnName.Contains('\0'))
            .Where(c => c.DataType == typeof(int) || c.DataType == typeof(long) ||
                c.DataType == typeof(short) || c.DataType == typeof(byte) ||
                c.DataType == typeof(bool) || c.DataType == typeof(DateTime) ||
                c.DataType == typeof(decimal) || c.DataType == typeof(double) ||
                c.DataType == typeof(float) || c.DataType == typeof(Guid) ||
                (c.DataType == typeof(string) && c.MaxLength is > 0 and <= 255))
            .ToArray();

        if (columns.Length == 0)
            throw new InvalidOperationException(
                "This table has no fields Access can safely compare to identify a record. " +
                "No records were deleted. Add a unique identifier in a copy of the database before deleting.");

        var parameters = new List<DataColumn>();
        var predicates = new List<string>();
        foreach (var column in columns)
        {
            string field = "[" + column.ColumnName + "]";
            object value = row[column, DataRowVersion.Original];
            if (value is null or DBNull)
                predicates.Add(field + " IS NULL");
            else
            {
                predicates.Add(field + " = ?");
                parameters.Add(column);
            }
        }
        return new KeylessMatch(string.Join(" AND ", predicates), parameters, columns);
    }

    internal static bool SameKeylessValues(DataRow left, DataRow right, IReadOnlyList<DataColumn> columns) =>
        columns.All(c => Equals(left[c, DataRowVersion.Original], right[c, DataRowVersion.Original]));

    public static int Stage(DataTable table, IEnumerable<DataRow> selected)
    {
        var records = EligibleRows(selected);
        if (records.Any(row => row.Table != table))
            throw new InvalidOperationException("Selected records belong to a different table.");

        foreach (var row in records)
        {
            // Newly inserted rows do not yet exist in the MDB. Remove them from
            // the local DataTable instead of marking them for a SQL DELETE.
            if (row.RowState == DataRowState.Added)
                table.Rows.Remove(row);
            else
                row.Delete();
        }

        // Existing rows remain pending until the user chooses Save Changes.
        return records.Count;
    }
}
