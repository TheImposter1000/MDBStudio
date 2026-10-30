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

    /// <summary>
    /// Copies data into NEW rows. When pasted keys or Recordname values are
    /// already used, generate fresh values using the full-table catalog.
    /// All rows remain unsaved until the explicit Save Changes operation.
    /// </summary>
    public static PasteResult Paste(DataTable table, PasteIdentityCatalog catalog,
        bool duplicateSelected, IReadOnlyList<DataRow>? selected = null)
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
            throw new InvalidOperationException("Paste at most 10,000 records at a time.");

        var columns = table.Columns.Cast<DataColumn>().ToArray();
        var fieldMap = fields.Select(field => columns.FirstOrDefault(c =>
            string.Equals(c.ColumnName, field, StringComparison.OrdinalIgnoreCase))).ToArray();
        if (fieldMap.All(c => c is null))
            throw new InvalidOperationException("No clipboard columns match this table.");

        var keys = table.PrimaryKey;
        var recordName = PasteIdentityCatalog.FindRecordName(table);
        var staged = new List<DataRow>();
        int generatedIds = 0;
        int generatedNames = 0;

        foreach (var values in sourceRows)
        {
            var row = table.NewRow();
            for (int i = 0; i < Math.Min(values.Length, fieldMap.Length); i++)
            {
                var column = fieldMap[i];
                if (column is null || column.AutoIncrement || column.ReadOnly ||
                    !string.IsNullOrEmpty(column.Expression)) continue;

                object? value = values[i];
                if (value is null || value == DBNull.Value ||
                    (value is string t && t.Length == 0))
                {
                    row[column] = column.DataType == typeof(string) && !column.AllowDBNull ?
                        "" : DBNull.Value;
                    continue;
                }
                try { row[column] = ConvertField(value, column.DataType); }
                catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException or ArgumentException)
                {
                    throw new InvalidOperationException(
                        $"Invalid value for '{column.ColumnName}' in pasted record {staged.Count + 1}: {ex.Message}", ex);
                }
            }

            // TEW-style Recordname fields often contain sequential text IDs such
            // as Prairie Cup013. Preserve the prefix and numbering width.
            if (recordName is not null && !row.IsNull(recordName) &&
                !recordName.ReadOnly && recordName.DataType == typeof(string))
            {
                string name = Convert.ToString(row[recordName], CultureInfo.InvariantCulture) ?? "";
                if (catalog.UsedRecordNames.Contains(name))
                {
                    row[recordName] = catalog.GenerateText(name, true, recordName.MaxLength);
                    generatedNames++;
                }
            }

            if (keys.Any(k => !k.AutoIncrement && row.IsNull(k)))
                throw new InvalidOperationException(
                    "The pasted records have missing primary keys. Supply the required IDs before pasting.");

            string key = PasteIdentityCatalog.KeyOf(row, keys);
            if (catalog.UsedKeys.Contains(key))
            {
                if (keys.Length != 1 || keys[0].AutoIncrement || keys[0].ReadOnly)
                    throw new InvalidOperationException(
                        "This record has a duplicate composite or protected primary key. " +
                        "MDB Studio cannot safely generate a replacement for this table structure.");

                var identity = keys[0];
                var type = identity.DataType;
                if (type == typeof(int) || type == typeof(long) ||
                    type == typeof(short) || type == typeof(byte))
                {
                    row[identity] = catalog.GenerateNumber(type);
                }
                else if (type == typeof(Guid))
                {
                    row[identity] = Guid.NewGuid();
                }
                else if (type == typeof(string))
                {
                    string original = Convert.ToString(row[identity], CultureInfo.InvariantCulture) ?? "";
                    bool isRecordName = recordName == identity;
                    row[identity] = catalog.GenerateText(original, isRecordName, identity.MaxLength);
                    if (isRecordName) generatedNames++;
                }
                else
                {
                    throw new InvalidOperationException(
                        $"MDB Studio cannot safely generate a primary key of type {type.Name}. " +
                        "Provide a unique primary key before pasting.");
                }
                generatedIds++;
                key = PasteIdentityCatalog.KeyOf(row, keys);
                if (catalog.UsedKeys.Contains(key))
                    throw new InvalidOperationException(
                        "Could not allocate a free record ID. No records were added.");
            }

            catalog.Include(keys.Select(k => row[k]).ToArray(),
                recordName is null || row.IsNull(recordName) ? null :
                Convert.ToString(row[recordName], CultureInfo.InvariantCulture));
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
            return new PasteResult(inserted.Count, generatedIds, generatedNames);
        }
        catch
        {
            foreach (var row in inserted) table.Rows.Remove(row);
            throw;
        }
    }

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

internal sealed record PasteResult(int Added, int GeneratedIds, int GeneratedNames);
