using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MDBStudio;

/// <summary>
/// Table-wide identity snapshot, including unsaved rows. The Access engine's
/// unique constraints remain the final authority when Save Changes runs.
/// </summary>
internal sealed class PasteIdentityCatalog
{
    public HashSet<string> UsedKeys { get; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> UsedSingleKeyValues { get; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> UsedRecordNames { get; } = new(StringComparer.OrdinalIgnoreCase);
    public long HighestNumericKey { get; private set; }
    private readonly Dictionary<string, long> _highestSuffix = new(StringComparer.OrdinalIgnoreCase);

    public void Include(object?[] keyValues, string? recordName)
    {
        UsedKeys.Add(KeyOf(keyValues));
        if (keyValues.Length == 1)
        {
            var value = Normalize(keyValues[0]);
            UsedSingleKeyValues.Add(value);
            if (keyValues[0] is not null and not DBNull &&
                long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number))
                HighestNumericKey = Math.Max(HighestNumericKey, number);
        }
        if (!string.IsNullOrEmpty(recordName)) UsedRecordNames.Add(recordName);
    }

    public void IncludePending(DataTable table)
    {
        var nameColumn = FindRecordName(table);
        foreach (DataRow row in table.Rows)
        {
            if (row.RowState == DataRowState.Deleted) continue;
            Include(table.PrimaryKey.Select(key => row[key]).ToArray(),
                nameColumn is null || row.IsNull(nameColumn) ? null :
                Convert.ToString(row[nameColumn], CultureInfo.InvariantCulture));
        }
    }

    public static DataColumn? FindRecordName(DataTable table) =>
        table.Columns.Cast<DataColumn>().FirstOrDefault(c =>
            string.Equals(c.ColumnName, "Recordname", StringComparison.OrdinalIgnoreCase));

    public static string KeyOf(DataRow row, DataColumn[] keys) =>
        KeyOf(keys.Select(key => row[key]).ToArray());

    public static string KeyOf(object?[] values) => string.Concat(values.Select(value =>
    {
        string text = Normalize(value);
        return text.Length.ToString(CultureInfo.InvariantCulture) + ":" + text + "|";
    }));

    private static string Normalize(object? value) =>
        value is null or DBNull ? "<null>" :
        Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";

    public object GenerateNumber(Type type)
    {
        if (HighestNumericKey == long.MaxValue)
            throw new InvalidOperationException("No available numeric IDs remain.");
        var candidate = checked(HighestNumericKey + 1);
        object result;
        try { result = Convert.ChangeType(candidate, type, CultureInfo.InvariantCulture); }
        catch (Exception ex) when (ex is OverflowException or InvalidCastException)
        {
            throw new InvalidOperationException(
                "The next record ID does not fit this column's numeric type. Enter a new ID manually.", ex);
        }
        HighestNumericKey = candidate;
        return result;
    }

    public string GenerateText(string original, bool recordName, int maxLength)
    {
        var used = recordName ? UsedRecordNames : UsedSingleKeyValues;
        var match = Regex.Match(original, @"^(.*?)(\d+)$", RegexOptions.CultureInvariant);
        string prefix;
        int width;
        if (match.Success)
        {
            prefix = match.Groups[1].Value;
            width = match.Groups[2].Value.Length;
        }
        else
        {
            prefix = original + " Copy ";
            width = 0;
        }

        string cacheKey = (recordName ? "N:" : "K:") + prefix;
        if (!_highestSuffix.TryGetValue(cacheKey, out long highest))
        {
            highest = 0;
            foreach (var candidate in used)
            {
                if (!candidate.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
                string suffix = candidate[prefix.Length..];
                if (long.TryParse(suffix, NumberStyles.None, CultureInfo.InvariantCulture, out var n))
                    highest = Math.Max(highest, n);
            }
        }

        string next;
        do
        {
            if (highest == long.MaxValue)
                throw new InvalidOperationException("No available numbered names remain.");
            highest++;
            next = prefix + highest.ToString(width > 0 ? "D" + width : "D", CultureInfo.InvariantCulture);
            if (maxLength > 0 && next.Length > maxLength)
                throw new InvalidOperationException(
                    $"Generated name '{next}' exceeds the maximum field length {maxLength}. Edit the identifier manually.");
        } while (used.Contains(next));

        _highestSuffix[cacheKey] = highest;
        return next;
    }
}
