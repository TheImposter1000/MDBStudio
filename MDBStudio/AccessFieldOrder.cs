namespace MDBStudio;

/// <summary>
/// Checks that a design-view field reorder is a permutation of the existing
/// fields. Reordering must never add, remove, rename or change a column.
/// </summary>
public static class AccessFieldOrder
{
    public static IReadOnlyList<string> Validate(
        IReadOnlyList<string> actual, IReadOnlyList<string> proposed)
    {
        if (actual.Count == 0)
            throw new InvalidOperationException("The selected table has no fields.");
        if (actual.Count != proposed.Count)
            throw new InvalidOperationException(
                "Field ordering cannot add or remove fields. Reload Design View and try again.");

        var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string name in actual)
        {
            AccessSchemaSql.Name(name);
            if (!names.TryAdd(name, name))
                throw new InvalidOperationException("Duplicate field in saved table: " + name);
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>(proposed.Count);
        foreach (string name in proposed)
        {
            if (string.IsNullOrEmpty(name) || !names.TryGetValue(name, out string? canonical) ||
                !seen.Add(name))
                throw new InvalidOperationException(
                    "The requested field order is missing, repeats or renames a field. " +
                    "Reload Design View and try again.");
            result.Add(canonical);
        }
        return result.ToArray();
    }

    public static bool Same(IReadOnlyList<string> first, IReadOnlyList<string> second) =>
        first.SequenceEqual(second, StringComparer.OrdinalIgnoreCase);
}
