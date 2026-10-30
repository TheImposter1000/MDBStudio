using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.VisualBasic.FileIO;

namespace MDBStudio;

/// <summary>
/// Copy rows to an ordinary spreadsheet-compatible TSV clipboard and preserve
/// typed values in memory while the clipboard is unchanged.
/// Never writes to the database. Pasted rows remain pending until Save Changes.
/// </summary>
internal static class RecordClipboard
{
    private sealed record Snapshot(string Text, string[] Fields, object?[][] Rows);
    private static Snapshot? _snapshot;

    public static int Copy(DataTable table, IReadOnlyList<DataRow> rows)
    {
        if (rows.Count == 0) return 0;
        var fields = table.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
        var copied = rows.Select(r => fields.Select(f =>
        {
            var value = r[f];
            return value is byte[] binary ? binary.ToArray() : value;
        }).ToArray()).ToArray();

        var text = new StringBuilder();
        text.AppendLine(string.Join('\t', fields.Select(TsvCell)));
        foreach (var row in copied)
        {
            text.AppendLine(string.Join('\t', row.Select(v => TsvCell(
                v is DBNull ? "" : v is byte[] bytes ? Convert.ToBase64String(bytes) :
                Convert.ToString(v, CultureInfo.CurrentCulture) ?? ""))));
        }
        string clipboard = text.ToString();
        Clipboard.SetText(clipboard);
        _snapshot = new Snapshot(clipboard, fields, copied);
        return copied.Length;
    }

    private static string TsvCell(string? value)
    {
        string text = value ?? "";
        return text.IndexOfAny(new[] {'\t', '\r', '\n', '"'}) >= 0
            ? "\"" + text.Replace("\"", "\"\"") + "\"" : text;
    }

    private static (string[] Fields, object?[][] Rows) ReadClipboard(DataTable table)
    {
        if (!Clipboard.ContainsText())
            throw new InvalidOperationException("The clipboard has no text records.");

        string text = Clipboard.GetText();
        if (_snapshot is not null && string.Equals(text, _snapshot.Text, StringComparison.Ordinal))
            return (_snapshot.Fields, _snapshot.Rows);

        using var parser = new TextFieldParser(new StringReader(text));
        parser.TextFieldType = FieldType.Delimited;
        parser.SetDelimiters("\t");
        parser.HasFieldsEnclosedInQuotes = true;
        parser.TrimWhiteSpace = false;

        var records = new List<string[]>();
        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields();
            if (fields is not null && fields.Length > 0)
                records.Add(fields);
            if (records.Count > 10000)
                throw new InvalidOperationException("Paste at most 10,000 records at a time.");
        }
        if (records.Count == 0)
            throw new InvalidOperationException("The clipboard has no records to paste.");

        var columns = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
        bool isHeader = records[0].Length > 0 &&
            records[0].All(f => columns.Contains(f, StringComparer.OrdinalIgnoreCase));
        var fieldNames = isHeader ? records[0] : columns;
        var rows = records.Skip(isHeader ? 1 : 0)
            .Where(row => row.Any(value => value.Length > 0))
            .Select(row => row.Cast<object?>().ToArray()).ToArray();
        if (rows.Length == 0)
            throw new InvalidOperationException("No clipboard data rows were found.");
        if (rows.Any(row => row.Length != fieldNames.Length))
            throw new InvalidOperationException("The pasted table has mismatched field counts. Copy entire rows or include a matching header.");

        return (fieldNames, rows);
    }

    public static int Paste(DataTable table, bool duplicateSelected, IReadOnlyList<DataRow>? selected = null)
    {
        if (table.PrimaryKey.Length == 0)
            throw new InvalidOperationException("This table cannot be edited without a primary key.");

        string[] fields;
        object?[][] sourceRows;
        if (duplicateSelected)
        {
            if (selected is null || selected.Count == 0)
                throw new InvalidOperationException("Select at least one row to duplicate.");
            fields = table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();
            sourceRows = selected.Select(r => fields.Select(field => r[field]).ToArray()).ToArray();
        }
        else
        {
            (fields, sourceRows) = ReadClipboard(table);
        }
        if (sourceRows.Length > 10000)
            throw new InvalidOperationException("Paste at most 10,000 rows at a time.");

        var keys = table.PrimaryKey;
        if (duplicateSelected && keys.Any(k => !k.AutoIncrement))
            throw new InvalidOperationException(
                "This table uses manually assigned primary keys. Copy the rows into a spreadsheet, " +
                "provide new unique IDs, then paste them back. Duplicating an existing ID is unsafe.");

        var fieldMap = fields.Select(name => table.Columns.Cast<DataColumn>()
            .FirstOrDefault(column => string.Equals(column.ColumnName, name, StringComparison.OrdinalIgnoreCase)))
            .ToArray();
        if (fieldMap.All(column => column is null))
            throw new InvalidOperationException("No clipboard column names match the destination table.");

        var existingKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (DataRow row in table.Rows)
        {
            if (row.RowState != DataRowState.Deleted)
                existingKeys.Add(KeyOf(row, keys));
        }

        var staged = new List<DataRow>();
        foreach (var values in sourceRows)
        {
            var row = table.NewRow();
            for (int i = 0; i < Math.Min(values.Length, fieldMap.Length); i++)
            {
                DataColumn? col = fieldMap[i];
                if (col is null || col.AutoIncrement || col.ReadOnly || !string.IsNullOrEmpty(col.Expression))
                    continue;
                object? value = values[i];
                if (value is null || value == DBNull.Value || (value is string t && t.Length == 0))
                {
                    row[col] = col.DataType == typeof(string) && !col.AllowDBNull ? "" : DBNull.Value;
                    continue;
                }

                try { row[col] = ConvertField(value, col.DataType); }
                catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or ArgumentException)
                {
                    throw new InvalidOperationException(
                        $"Invalid value for column '{col.ColumnName}' in pasted record {staged.Count + 1}: {ex.Message}", ex);
                }
            }

            if (keys.Any(k => !k.AutoIncrement && row.IsNull(k)))
                throw new InvalidOperationException(
                    "Pasted records require values for all manually assigned primary keys.");
            string key = KeyOf(row, keys);
            if (!existingKeys.Add(key))
                throw new InvalidOperationException("Paste would create a duplicate primary key. No records were added.");
            staged.Add(row);
        }

        var inserted = new List<DataRow>();
        try
        {
            foreach (var row in staged)
            {
                table.Rows.Add(row);
                inserted.Add(row);
            }
            return inserted.Count;
        }
        catch
        {
            foreach (var row in inserted) table.Rows.Remove(row);
            throw;
        }
    }

    private static string KeyOf(DataRow row, DataColumn[] keys) =>
        string.Join("\u001f", keys.Select(k => row[k] == DBNull.Value ? "<null>" :
            Convert.ToString(row[k], CultureInfo.InvariantCulture) ?? ""));

    private static object ConvertField(object value, Type type)
    {
        if (type.IsInstanceOfType(value)) return value;
        var text = Convert.ToString(value, CultureInfo.CurrentCulture) ?? "";
        if (type == typeof(Guid)) return Guid.Parse(text);
        if (type == typeof(byte[])) return Convert.FromBase64String(text);
        if (type == typeof(bool))
        {
            if (text == "1") return true;
            if (text == "0") return false;
            return bool.Parse(text);
        }
        if (type == typeof(DateTime)) return DateTime.Parse(text, CultureInfo.CurrentCulture);
        return Convert.ChangeType(value, type, CultureInfo.CurrentCulture);
    }
}
