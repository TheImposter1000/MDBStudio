using ClosedXML.Excel;
using System.Data;
using System.Globalization;

namespace MDBStudio;

/// <summary>
/// Excel transfer for MDB Studio. No Office/Excel installation is needed.
/// Imports worksheet values as text so IDs with leading zeros are preserved.
/// </summary>
public static class ExcelTransfer
{
    public const int MaxImportRows = 100_000;
    public const int MaxColumns = 200;
    public const int MaxFileBytes = 50 * 1024 * 1024;
    private const int ExcelMaxRows = 1_048_576;

    public static IReadOnlyList<string> Sheets(string filename)
    {
        CheckFile(filename);
        using var book = new XLWorkbook(filename);
        return book.Worksheets.Select(s => s.Name).ToArray();
    }

    private static void CheckFile(string filename)
    {
        if (!string.Equals(Path.GetExtension(filename), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Please select an Excel .xlsx workbook. Legacy .xls files are not supported.");
        var info = new FileInfo(filename);
        if (!info.Exists || info.Length > MaxFileBytes)
            throw new InvalidOperationException("Excel workbook is missing or exceeds 50 MB.");
    }

    public static DataTable ReadSheet(string filename, string sheetName)
    {
        CheckFile(filename);
        using var book = new XLWorkbook(filename);
        var sheet = book.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
            ?? throw new InvalidOperationException("The selected worksheet is no longer in the workbook.");
        int headerRow = sheet.FirstRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
        int lastRow = sheet.LastRowUsed(XLCellsUsedOptions.Contents)?.RowNumber() ?? 0;
        int lastColumn = sheet.LastColumnUsed(XLCellsUsedOptions.Contents)?.ColumnNumber() ?? 0;
        if (headerRow == 0 || lastRow <= headerRow || lastColumn == 0)
            throw new InvalidOperationException("The worksheet must contain a header row and at least one data row.");
        if (lastColumn > MaxColumns)
            throw new InvalidOperationException($"Excel import supports a maximum of {MaxColumns} columns.");
        if (lastRow - headerRow > MaxImportRows)
            throw new InvalidOperationException($"Import at most {MaxImportRows:N0} Excel rows at a time. Split larger workbooks.");

        var table = new DataTable(sheet.Name) { Locale = CultureInfo.InvariantCulture };
        var reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int col = 1; col <= lastColumn; col++)
        {
            string raw = sheet.Cell(headerRow, col).GetString().Trim();
            string name = string.IsNullOrWhiteSpace(raw) ? $"Column{col}" : raw;
            name = name.Replace("[", "_").Replace("]", "_").Replace('\0', '_');
            if (name.Length > 58) name = name[..58];
            string original = name;
            int n = 2;
            while (!reserved.Add(name))
                name = original[..Math.Min(original.Length, 55)] + "_" + n++;
            table.Columns.Add(name, typeof(string));
        }

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var values = new object[lastColumn];
            bool hasContent = false;
            for (int c = 1; c <= lastColumn; c++)
            {
                IXLCell cell = sheet.Cell(r, c);
                string value;
                if (cell.IsEmpty()) value = "";
                else if (cell.DataType == XLDataType.DateTime)
                    value = cell.GetDateTime().ToString("O", CultureInfo.InvariantCulture);
                else if (cell.DataType == XLDataType.Boolean)
                    value = cell.GetBoolean() ? "True" : "False";
                else
                    value = cell.GetFormattedString();
                values[c - 1] = value;
                hasContent |= value.Length != 0;
            }
            if (hasContent) table.Rows.Add(values);
        }
        if (table.Rows.Count == 0)
            throw new InvalidOperationException("The worksheet has no nonblank records.");
        table.AcceptChanges();
        return table;
    }

    public static void ExportQuery(DataTable results, string filename, string sheetName = "Query Results")
    {
        if (results.Columns.Count == 0)
            throw new InvalidOperationException("Run a SELECT query before exporting to Excel.");
        if (results.Rows.Count >= ExcelMaxRows)
            throw new InvalidOperationException("Excel supports 1,048,575 data rows per sheet. Narrow the query.");
        if (results.Columns.Count > 16_384)
            throw new InvalidOperationException("Too many columns for an Excel worksheet.");
        using var book = new XLWorkbook();
        var ws = book.Worksheets.Add(SheetName(sheetName));
        int columnCount = results.Columns.Count;
        for (int c = 0; c < columnCount; c++)
            ws.Cell(1, c + 1).SetValue(results.Columns[c].ColumnName);

        int rowIndex = 2;
        foreach (DataRow row in results.Rows)
        {
            if (row.RowState == DataRowState.Deleted) continue;
            for (int c = 0; c < columnCount; c++)
            {
                var cell = ws.Cell(rowIndex, c + 1);
                var value = row[c];
                if (value is DBNull or null) continue;
                switch (value)
                {
                    case DateTime date: cell.SetValue(date); cell.Style.DateFormat.Format = "yyyy-mm-dd hh:mm:ss"; break;
                    case bool boolean: cell.SetValue(boolean); break;
                    case byte or short or int: cell.SetValue(Convert.ToInt32(value, CultureInfo.InvariantCulture)); break;
                    case long integer when Math.Abs((decimal)integer) <= 999_999_999_999_999M:
                        cell.SetValue((double)integer); break;
                    case long integer: cell.SetValue(integer.ToString(CultureInfo.InvariantCulture)); break;
                    case double number: cell.SetValue(number); break;
                    case float single: cell.SetValue((double)single); break;
                    case decimal money: cell.SetValue((double)money); break;
                    case byte[] blob: cell.SetValue(Convert.ToBase64String(blob)); break;
                    default: cell.SetValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""); break;
                }
            }
            rowIndex++;
        }

        var header = ws.Range(1, 1, 1, columnCount);
        header.Style.Font.Bold = true;
        header.Style.Font.FontColor = XLColor.White;
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#174A83");
        ws.SheetView.FreezeRows(1);
        if (rowIndex > 2)
            ws.Range(1, 1, rowIndex - 1, columnCount).SetAutoFilter();
        for (int c = 1; c <= columnCount; c++)
        {
            // Predictable widths for large exports without expensive auto-fit scans.
            ws.Column(c).Width = 21;
        }
        book.SaveAs(filename);
    }

    private static string SheetName(string name)
    {
        var safe = new string(name.Where(c => !"[]:*?/\\".Contains(c)).ToArray()).Trim();
        return string.IsNullOrEmpty(safe) ? "Query Results" : safe[..Math.Min(safe.Length, 31)];
    }
}
