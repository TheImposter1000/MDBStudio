using System.IO;
using System.Text.Json;

namespace MDBStudio;

/// <summary>
/// Saves the visible datasheet column order per database file and table.
/// This changes display position, not the Access schema or stored values.
/// </summary>
public static class TableColumnLayout
{
    private static string SettingsFile => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MDBStudio", "column-layouts.json");

    public static string[] Apply(IReadOnlyList<string> actual, IReadOnlyList<string> desired)
    {
        var existing = new HashSet<string>(actual, StringComparer.OrdinalIgnoreCase);
        var selected = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ordered = new List<string>(actual.Count);
        foreach (string name in desired)
            if (existing.Contains(name) && selected.Add(name))
                ordered.Add(actual.First(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)));
        foreach (string name in actual)
            if (selected.Add(name)) ordered.Add(name);
        return ordered.ToArray();
    }

    private static string Key(string filename, string table) =>
        Path.GetFullPath(filename) + " :: " + table;

    public static IReadOnlyList<string> Load(string filename, string table)
    {
        try
        {
            if (!File.Exists(SettingsFile)) return Array.Empty<string>();
            var content = JsonSerializer.Deserialize<Dictionary<string, string[]>>(
                File.ReadAllText(SettingsFile));
            if (content is null) return Array.Empty<string>();
            return new Dictionary<string, string[]>(content, StringComparer.OrdinalIgnoreCase)
                .GetValueOrDefault(Key(filename, table)) ?? Array.Empty<string>();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            return Array.Empty<string>();
        }
    }

    public static void Save(string filename, string table, IReadOnlyList<string> displayed)
    {
        try
        {
            var path = SettingsFile;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            var current = File.Exists(path)
                ? JsonSerializer.Deserialize<Dictionary<string, string[]>>(File.ReadAllText(path))
                : null;
            var entries = current is null
                ? new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
                : new Dictionary<string, string[]>(current, StringComparer.OrdinalIgnoreCase);
            entries[Key(filename, table)] = displayed.ToArray();
            File.WriteAllText(path, JsonSerializer.Serialize(entries));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            // Cosmetic layout persistence should never prevent database editing.
        }
    }
}
