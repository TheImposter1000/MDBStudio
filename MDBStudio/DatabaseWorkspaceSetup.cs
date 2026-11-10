namespace MDBStudio;

/// <summary>
/// Supplies the first editable table design for a newly created Access database.
/// This is a proposal shown to the user, never a silent schema change.
/// </summary>
public static class DatabaseWorkspaceSetup
{
    public static string SuggestTableName(IEnumerable<string> existingNames)
    {
        var existing = new HashSet<string>(existingNames, StringComparer.OrdinalIgnoreCase);
        for (int number = 1; number < int.MaxValue; number++)
        {
            string candidate = "Table" + number;
            if (!existing.Contains(candidate)) return candidate;
        }
        throw new InvalidOperationException("No available default table name was found.");
    }

    public static IReadOnlyList<AccessFieldDefinition> StarterFields() =>
        new[]
        {
            new AccessFieldDefinition("ID", AccessFieldKind.AutoNumber, PrimaryKey: true),
            new AccessFieldDefinition("Name", AccessFieldKind.ShortText, Length: 100)
        };
}
