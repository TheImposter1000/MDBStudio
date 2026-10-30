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

    public static int Stage(DataTable table, IEnumerable<DataRow> selected)
    {
        if (table.PrimaryKey.Length == 0)
            throw new InvalidOperationException(
                "Deleting records requires a detectable primary key.");

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
