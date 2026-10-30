using System.Data;
using System.Globalization;
using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace MDBStudio;

/// <summary>
/// Spreadsheet-style copying and pasting of VALUES into existing cells.
/// Does not add or duplicate database records.
/// </summary>
public static class GridCellEditor
{
    public sealed record CellValue(int RowIndex, int ColumnIndex, object Value);
    public sealed record CellPatch(DataRow Row, DataColumn Column, string Text);

    /// <summary>Copy an exact rectangular selection, with no record headers or IDs added.</summary>
    public static string CopyValues(IReadOnlyList<CellValue> selected)
    {
        if (selected.Count == 0) throw new InvalidOperationException("Select cells to copy.");
        var rows = selected.Select(c => c.RowIndex).Distinct().Order().ToArray();
        var columns = selected.Select(c => c.ColumnIndex).Distinct().Order().ToArray();
        if ((long)rows.Length * columns.Length != selected.Count ||
            rows[^1] - rows[0] + 1 != rows.Length ||
            columns[^1] - columns[0] + 1 != columns.Length)
            throw new InvalidOperationException(
                "Select one continuous rectangular range when copying multiple cells.");

        var map = selected.ToDictionary(c => (c.RowIndex, c.ColumnIndex), c => c.Value);
        var output = new StringBuilder();
        for (int r = 0; r < rows.Length; r++)
        {
            if (r != 0) output.Append("\r\n");
            for (int c = 0; c < columns.Length; c++)
            {
                if (c != 0) output.Append('\t');
                if (!map.TryGetValue((rows[r], columns[c]), out object? value))
                    throw new InvalidOperationException("The selected cell block is incomplete.");
                output.Append(Escape(ToText(value)));
            }
        }
        return output.ToString();
    }

    private static string ToText(object? value) => value switch
    {
        null or DBNull => "",
        DateTime date => date.ToString("O", CultureInfo.InvariantCulture),
        byte[] bytes => Convert.ToBase64String(bytes),
        _ => Convert.ToString(value, CultureInfo.CurrentCulture) ?? ""
    };

    private static string Escape(string value) =>
        value.IndexOfAny(new[] { '\t', '\r', '\n', '"' }) >= 0
            ? "\"" + value.Replace("\"", "\"\"") + "\"" : value;

    /// <summary>Read tab-separated clipboard values, including Excel quoted multi-line cells.</summary>
    public static string[][] ParseValues(string text)
    {
        if (string.IsNullOrEmpty(text)) return Array.Empty<string[]>();
        using var parser = new TextFieldParser(new StringReader(text));
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters("\t");
        parser.HasFieldsEnclosedInQuotes = true;
        parser.TrimWhiteSpace = false;
        var rows = new List<string[]>();
        while (!parser.EndOfData)
        {
            string[]? fields = parser.ReadFields();
            if (fields is null) continue;
            rows.Add(fields);
            if (rows.Count > 50_000)
                throw new InvalidOperationException("Paste a maximum of 50,000 rows at once.");
        }
        if (rows.Count == 0) return Array.Empty<string[]>();
        int columns = rows[0].Length;
        if (rows.Any(r => r.Length != columns))
            throw new InvalidOperationException("Clipboard data must have the same number of columns on every row.");
        if ((long)rows.Count * columns > 100_000)
            throw new InvalidOperationException("Paste a maximum of 100,000 cells at once.");
        return rows.ToArray();
    }

    /// <summary>
    /// Change only the selected fields of existing rows. Validate all cells
    /// before writing any values. The user must still click Save Changes.
    /// </summary>
    public static int Apply(IReadOnlyList<CellPatch> patches)
    {
        if (patches.Count == 0) return 0;
        if (patches.Count > 100_000)
            throw new InvalidOperationException("Modify at most 100,000 cells at once.");

        var staged = new List<(DataRow Row, DataColumn Column, object Value, object Old)>();
        var seen = new HashSet<(DataRow, DataColumn)>();
        foreach (var patch in patches)
        {
            DataRow row = patch.Row;
            DataColumn col = patch.Column;
            if (!seen.Add((row, col))) throw new InvalidOperationException("The destination contains repeated cells.");
            if (row.Table is null || col.Table != row.Table ||
                row.RowState is DataRowState.Deleted or DataRowState.Detached)
                throw new InvalidOperationException("The destination record is no longer available.");
            if (col.ReadOnly || col.AutoIncrement || !string.IsNullOrEmpty(col.Expression) ||
                row.Table.PrimaryKey.Contains(col))
                throw new InvalidOperationException(
                    $"'{col.ColumnName}' is a protected or primary key column. Select ordinary data cells instead.");
            object value;
            try { value = ConvertCell(patch.Text, col); }
            catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or ArgumentException)
            {
                throw new InvalidOperationException(
                    $"Cannot paste '{patch.Text}' into column '{col.ColumnName}' ({col.DataType.Name}): {ex.Message}", ex);
            }
            object old = row[col];
            if (!Equals(old, value)) staged.Add((row, col, value, old));
        }

        var applied = new List<(DataRow Row, DataColumn Column, object Old)>();
        try
        {
            foreach (var item in staged)
            {
                item.Row[item.Column] = item.Value;
                applied.Add((item.Row, item.Column, item.Old));
            }
        }
        catch
        {
            for (int i = applied.Count - 1; i >= 0; i--)
            {
                var prior = applied[i];
                try { prior.Row[prior.Column] = prior.Old; } catch { /* preserve underlying error */ }
            }
            throw;
        }
        return staged.Count;
    }

    private static object ConvertCell(string text, DataColumn col)
    {
        Type type = col.DataType;
        if (type == typeof(string))
        {
            if (col.MaxLength >= 0 && text.Length > col.MaxLength)
                throw new FormatException($"The value exceeds the {col.MaxLength}-character field limit.");
            return text;
        }
        if (text.Length == 0)
        {
            if (col.AllowDBNull) return DBNull.Value;
            throw new FormatException("This column cannot be blank.");
        }
        if (type == typeof(bool))
        {
            if (text == "1") return true;
            if (text == "0") return false;
            return bool.Parse(text);
        }
        if (type == typeof(Guid)) return Guid.Parse(text);
        if (type == typeof(DateTime))
        {
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime date) ||
                DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out date))
                return date;
            throw new FormatException("Invalid date.");
        }
        if (type == typeof(byte[]))
            throw new FormatException("Binary fields cannot be pasted in the datasheet.");
        if (type == typeof(double) || type == typeof(float) || type == typeof(decimal) ||
            type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte))
        {
            try { return Convert.ChangeType(text, type, CultureInfo.CurrentCulture); }
            catch (FormatException) { return Convert.ChangeType(text, type, CultureInfo.InvariantCulture); }
        }
        throw new InvalidOperationException($"Pasting is not supported for {type.Name} columns.");
    }
}
