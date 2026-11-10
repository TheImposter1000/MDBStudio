namespace MDBStudio;

/// <summary>
/// One secondary Access index that can be recreated on a replacement empty table.
/// The real primary key is defined by the proposed column definitions.
/// </summary>
public sealed record AccessRedesignIndex(
    string Name, IReadOnlyList<string> Columns, bool Unique);

public sealed record EmptyTableRedesignPlan(
    string TableName,
    IReadOnlyList<AccessFieldDefinition> Fields,
    IReadOnlyList<AccessRedesignIndex> RetainedIndexes,
    IReadOnlyList<string> RemovedIndexNames,
    string CreateTableSql)
{
    public string Description => RemovedIndexNames.Count == 0
        ? "All supported secondary indexes will be retained."
        : "The following indexes cannot be retained with this design and will be removed: " +
          string.Join(", ", RemovedIndexNames) + ".";
}

public static class EmptyTableRedesign
{
    /// <summary>
    /// Never transform a populated table by dropping it. Rebuilding is only
    /// permitted for verified empty tables. For empty tables, carry over
    /// secondary indexes whose columns still exist and support indexing.
    /// </summary>
    public static EmptyTableRedesignPlan Plan(
        string tableName, long existingRows,
        IReadOnlyList<AccessFieldDefinition> fields,
        IReadOnlyList<AccessRedesignIndex> secondaryIndexes)
    {
        if (existingRows != 0)
            throw new InvalidOperationException(
                "This table contains " + existingRows.ToString("N0") +
                " records. Full table redesign is only available for empty tables. " +
                "For populated tables, use Add field and Change field type for non-key fields. " +
                "MDB Studio will not delete or rebuild a table containing existing records.");

        string create = AccessSchemaSql.CreateTable(tableName, fields);
        var known = fields.ToDictionary(f => f.Name, f => f.Kind,
            StringComparer.OrdinalIgnoreCase);
        var keep = new List<AccessRedesignIndex>();
        var remove = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var index in secondaryIndexes)
        {
            AccessSchemaSql.Name(index.Name);
            if (!seen.Add(index.Name))
                throw new InvalidOperationException(
                    "The existing Access schema returned a duplicate index name: " + index.Name);
            if (index.Columns.Count == 0)
                throw new InvalidOperationException(
                    "Index " + index.Name + " has no readable fields; table redesign was cancelled.");
            foreach (string name in index.Columns) AccessSchemaSql.Name(name);
            bool supported = index.Columns.Count <= 10 &&
                index.Columns.All(name =>
                    known.TryGetValue(name, out var kind) &&
                    kind != AccessFieldKind.LongText);
            if (supported)
                keep.Add(index);
            else
                remove.Add(index.Name);
        }

        return new EmptyTableRedesignPlan(tableName, fields.ToArray(),
            keep.ToArray(), remove.ToArray(), create);
    }
}
