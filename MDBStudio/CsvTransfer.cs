using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace MDBStudio;

/// <summary>
/// CSV interchange for the spreadsheet import dialog and SQL query results.
/// Reads quoted commas, double quotes and multiline cells; never executes
/// spreadsheet formulas. The first CSV record is the header row.
/// </summary>
public static class CsvTransfer
{
    public const int MaxImportRows = 100_000;
    public const int MaxColumns = 200;
    public const long MaxFileBytes = 50L * 1024 * 1024;
    public const string FileLabel = "CSV data";

    public static DataTable ReadFile(string filename)
    {
        if (!string.Equals(Path.GetExtension(filename), ".csv", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Select a .csv file.");
        var info = new FileInfo(filename);
        if (!info.Exists || info.Length > MaxFileBytes)
            throw new InvalidOperationException("CSV file is missing or exceeds the 50 MB limit.");

        // TextFieldParser implements the standard quoted-field CSV rules,
        // including embedded delimiters, newlines and escaped quotes.
        using var parser = new TextFieldParser(filename, Encoding.UTF8, detectEncoding: true)
        {
            TextFieldType = FieldType.Delimited,
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false
        };
        parser.SetDelimiters(",");
        string[]? headers;
        try { headers = parser.ReadFields(); }
        catch (MalformedLineException ex)
        {
            throw new InvalidOperationException("Invalid CSV header row: " + ex.Message, ex);
        }
        if (headers is null || headers.Length == 0 || headers.All(string.IsNullOrWhiteSpace))
            throw new InvalidOperationException("The CSV file must start with a nonblank column header row.");
        if (headers.Length > MaxColumns)
            throw new InvalidOperationException($"CSV import supports no more than {MaxColumns} columns.");

        var table = new DataTable(Path.GetFileNameWithoutExtension(filename))
        {
            Locale = CultureInfo.InvariantCulture
        };
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int c = 0; c < headers.Length; c++)
        {
            string label = (headers[c] ?? "").Trim().TrimStart('\uFEFF')
                .Replace("[", "_").Replace("]", "_").Replace('\0', '_');
            if (string.IsNullOrWhiteSpace(label)) label = $"Column{c + 1}";
            if (label.Length > 58) label = label[..58];
            string origin = label;
            int suffix = 2;
            while (!names.Add(label))
                label = origin[..Math.Min(55, origin.Length)] + "_" + suffix++;
            table.Columns.Add(label, typeof(string));
        }

        int record = 1;
        while (!parser.EndOfData)
        {
            string[]? fields;
            try { fields = parser.ReadFields(); }
            catch (MalformedLineException ex)
            {
                throw new InvalidOperationException(
                    $"Malformed CSV record around line {parser.ErrorLineNumber}: {ex.Message}", ex);
            }
            record++;
            if (fields is null) continue;
            if (fields.Length != table.Columns.Count)
                throw new InvalidOperationException(
                    $"CSV data record {record} has {fields.Length} fields; expected {table.Columns.Count}. " +
                    "Ensure all fields containing commas or newlines are correctly quoted.");
            if (fields.All(value => string.IsNullOrEmpty(value))) continue;
            table.Rows.Add(fields.Cast<object>().ToArray());
            if (table.Rows.Count > MaxImportRows)
                throw new InvalidOperationException(
                    $"Import at most {MaxImportRows:N0} records from one CSV file.");
        }
        if (table.Rows.Count == 0)
            throw new InvalidOperationException("The CSV file contains headers but no data rows.");
        table.AcceptChanges();
        return table;
    }

    public static void ExportQuery(DataTable result, string filename)
    {
        if (result.Columns.Count == 0)
            throw new InvalidOperationException("Run a SELECT query before exporting to CSV.");

        // Write UTF-8 with BOM, which Excel recognizes as UTF-8 when opening CSV.
        // The full query result is exported, not the currently selected cells.
        using var stream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        WriteRecord(writer, result.Columns.Cast<DataColumn>().Select(c => c.ColumnName), escapeFormulas: false);
        foreach (DataRow row in result.Rows)
        {
            if (row.RowState == DataRowState.Deleted) continue;
            var cells = result.Columns.Cast<DataColumn>().Select(column =>
            {
                object value = row[column];
                if (value is DBNull or null) return (Text: "", Guard: false);
                return (Text: value switch
                {
                    DateTime date => date.ToString("O", CultureInfo.InvariantCulture),
                    bool boolean => boolean ? "True" : "False",
                    byte[] bytes => Convert.ToBase64String(bytes),
                    _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
                }, Guard: column.DataType == typeof(string));
            }).ToArray();
            WriteRecord(writer, cells);
        }
    }

    private static void WriteRecord(TextWriter writer, IEnumerable<string> values, bool escapeFormulas)
    {
        WriteRecord(writer, values.Select(v => (Text: v, Guard: escapeFormulas)));
    }

    private static void WriteRecord(TextWriter writer, IEnumerable<(string Text, bool Guard)> cells)
    {
        bool first = true;
        foreach (var item in cells)
        {
            if (!first) writer.Write(',');
            first = false;
            string value = item.Text;
            // CSV is often opened directly by spreadsheet applications.
            // Prefix potentially executable text with an apostrophe to stop
            // formula evaluation; original MDB values remain unchanged.
            if (item.Guard && IsFormulaLike(value)) value = "'" + value;
            writer.Write('"');
            writer.Write(value.Replace("\"", "\"\""));
            writer.Write('"');
        }
        writer.Write("\r\n");
    }

    private static bool IsFormulaLike(string value)
    {
        int i = 0;
        while (i < value.Length && (value[i] == ' ' || value[i] == '\t' ||
                                    value[i] == '\r' || value[i] == '\n')) i++;
        return i < value.Length && "=+-@".Contains(value[i]);
    }
}
