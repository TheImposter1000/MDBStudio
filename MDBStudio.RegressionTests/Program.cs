using System.IO;
using System.Data;
using MDBStudio;

internal static class Program
{
    [STAThread]
    public static int Main()
    {
        TestSequentialIdsAndNames();
        TestRepeatedPaste();
        TestCompositeKeyIsProtected();
        TestGuidKey();
        TestCellCopyAndBroadcast();
        TestMultiColumnPaste();
        TestProtectedKeysAndInvalidValues();
        TestTextClipboardRoundtrip();
        TestDeleteSelectedRecords();
        TestDeletePendingInsertedRecord();
        TestDeleteWithoutPrimaryKey();
        TestKeylessDeletePredicateSafety();
        TestExcelQueryExportAndImport();
        TestExcelFileRejections();
        TestCsvQuotedImportAndExport();
        TestCsvMalformedRows();
        TestCsvFormulaEscaping();
        TestAccessBooleanSpreadsheetValues();
        TestAccessBooleanCsvImport();
        TestAccessBooleanCellPaste();
        TestNumericClipboardAndLegacyCsvEscapes();
        TestSignedNumericCsvExportRoundtrip();
        TestNumericImportDiagnostics();
        TestColumnFilters();
        Console.WriteLine("Paste identity regression tests passed.");
        return 0;
    }

    private static void TestColumnFilters()
    {
        var table = new DataTable("tblShowHistory");
        var company = table.Columns.Add("Company", typeof(string));
        company.MaxLength = 80;
        var name = table.Columns.Add("WorkerName", typeof(string));
        name.MaxLength = 100;
        var winner = table.Columns.Add("Victor", typeof(int));
        var showDate = table.Columns.Add("EventDate", typeof(DateTime));
        table.Columns.Add("Active", typeof(bool));
        table.Columns.Add("MemoBlob", typeof(byte[]));

        var criteria = new[]
        {
            new TableColumnFilter("Company", TableFilterOperation.Contains, "DNA"),
            new TableColumnFilter("WorkerName", TableFilterOperation.StartsWith, "Jayden"),
            new TableColumnFilter("Victor", TableFilterOperation.Equals, "-1"),
            new TableColumnFilter("EventDate", TableFilterOperation.IsBlank, "")
        };
        var query = TableFilters.Compile(table.Columns, criteria);
        Check(query.WhereClause.Contains("InStr(1, [Company], ?, 1) > 0"),
            "Contains must use a parameterized case-insensitive Access search.");
        Check(query.WhereClause.Contains("InStr(1, [WorkerName], ?, 1) = 1"),
            "Starts with must search the whole source table.");
        Check(query.WhereClause.Contains("[Victor] = ?") &&
              query.WhereClause.Contains("[EventDate] IS NULL"),
            "Numeric and missing-value filters must have correct SQL conditions.");
        Check(query.Parameters.Count == 3 &&
              (string)query.Parameters[0].Value == "DNA" &&
              (string)query.Parameters[1].Value == "Jayden" &&
              (int)query.Parameters[2].Value == -1,
            "Parameter types and order must be preserved, including TEW Victor=-1.");

        var textAttack = TableFilters.Compile(table.Columns, new[]
        {
            new TableColumnFilter("Company", TableFilterOperation.Contains,
                "' OR 1=1 --")
        });
        Check(!textAttack.WhereClause.Contains("1=1") &&
              (string)textAttack.Parameters[0].Value == "' OR 1=1 --",
            "Filter values must be parameters, not SQL fragments.");

        var suffix = TableFilters.Compile(table.Columns, new[]
        {
            new TableColumnFilter("WorkerName", TableFilterOperation.EndsWith, "Cole")
        });
        Check(suffix.Parameters.Count == 2 &&
              suffix.WhereClause.Contains("Right([WorkerName], Len(?)) = ?"),
            "Ends with must bind both positional parameters.");

        var blank = TableFilters.Compile(table.Columns, new[]
        {
            new TableColumnFilter("Company", TableFilterOperation.IsBlank, "")
        });
        Check(blank.Parameters.Count == 0 &&
              blank.WhereClause.Contains("[Company] = ''"),
            "Blank text must match both NULL and empty text.");

        try
        {
            TableFilters.Compile(table.Columns, new[]
            {
                new TableColumnFilter("Victor", TableFilterOperation.Contains, "1")
            });
            throw new Exception("Numeric columns must reject text-only operators.");
        }
        catch (InvalidOperationException) { }

        try
        {
            TableFilters.Compile(table.Columns, new[]
            {
                new TableColumnFilter("MissingColumn", TableFilterOperation.Equals, "x")
            });
            throw new Exception("Unknown columns must not be injected into SQL.");
        }
        catch (InvalidOperationException) { }

        try
        {
            TableFilters.Compile(table.Columns, new[]
            {
                new TableColumnFilter("MemoBlob", TableFilterOperation.Equals, "data")
            });
            throw new Exception("Binary columns must not be compared by SQL filters.");
        }
        catch (InvalidOperationException) { }

        try
        {
            TableFilters.Compile(table.Columns, new[]
            {
                new TableColumnFilter("Company", TableFilterOperation.Equals, "DNA"),
                new TableColumnFilter("Company", TableFilterOperation.Equals, "NWE")
            });
            throw new Exception("Duplicate column filters should not be silently applied.");
        }
        catch (InvalidOperationException) { }

        Check(TableFilters.Compile(table.Columns, Array.Empty<TableColumnFilter>())
              .WhereClause == "", "No active filters must return every database record.");
    }

    private static DataTable Example()
    {
        var table = new DataTable("tblBeltPreHistory");
        var id = table.Columns.Add("UID", typeof(int));
        table.Columns.Add("Recordname", typeof(string));
        table.Columns.Add("BeltUID", typeof(int));
        table.Columns.Add("BeltOrder", typeof(int));
        table.PrimaryKey = new[] { id };
        table.Rows.Add(13, "Prairie Cup013", 0, 13);
        table.Rows.Add(14, "Prairie Cup014", 0, 14);
        table.AcceptChanges();
        return table;
    }

    private static PasteIdentityCatalog Snapshot(DataTable table)
    {
        var identities = new PasteIdentityCatalog();
        identities.Include(new object?[] { 9000 }, "Prairie Cup023");
        identities.IncludePending(table);
        return identities;
    }

    private static void TestSequentialIdsAndNames()
    {
        var table = Example();
        var selected = table.Rows.Cast<DataRow>().ToArray();
        var result = RecordClipboard.Paste(table, Snapshot(table), true, selected);
        Check(result.Added == 2 && result.GeneratedIds == 2 && result.GeneratedNames == 2,
            "Two new IDs and numbered record names must be allocated.");

        var added = table.Rows.Cast<DataRow>().Where(r => r.RowState == DataRowState.Added).ToArray();
        Check((int)added[0]["UID"] == 9001 && (int)added[1]["UID"] == 9002, "Numeric keys must use global maximum.");
        Check((string)added[0]["Recordname"] == "Prairie Cup024" &&
              (string)added[1]["Recordname"] == "Prairie Cup025", "Numbered names must preserve prefix and padding.");
        Check((int)added[0]["BeltOrder"] == 13 && (int)added[1]["BeltOrder"] == 14,
            "Unrelated fields must not be changed.");
    }

    private static void TestRepeatedPaste()
    {
        var table = Example();
        var selected = table.Rows.Cast<DataRow>().ToArray();
        RecordClipboard.Paste(table, Snapshot(table), true, selected);
        var second = RecordClipboard.Paste(table, Snapshot(table), true, selected);
        Check(second.Added == 2, "Repeated paste must produce another set of records.");
        var added = table.Rows.Cast<DataRow>().Where(r => r.RowState == DataRowState.Added).ToArray();
        Check((int)added[2]["UID"] == 9003, "Unsaved IDs must also be reserved.");
        Check((string)added[2]["Recordname"] == "Prairie Cup026", "Unsaved names must also be reserved.");
    }

    private static void TestCompositeKeyIsProtected()
    {
        var table = Example();
        table.PrimaryKey = new[] { table.Columns["UID"]!, table.Columns["BeltUID"]! };
        var before = table.Rows.Count;
        try
        {
            RecordClipboard.Paste(table, Snapshot(table), true, new[] { table.Rows[0] });
            throw new Exception("Expected unsupported composite key paste to be rejected.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("composite"))
        {
            Check(table.Rows.Count == before, "Failed paste must not insert records.");
        }
    }

    private static void TestGuidKey()
    {
        var table = new DataTable("GuidTable");
        var id = table.Columns.Add("UID", typeof(Guid));
        table.Columns.Add("Value", typeof(string));
        table.PrimaryKey = new[] { id };
        var old = Guid.NewGuid();
        table.Rows.Add(old, "Data");
        table.AcceptChanges();
        var identities = new PasteIdentityCatalog();
        identities.IncludePending(table);
        var result = RecordClipboard.Paste(table, identities, true, new[] { table.Rows[0] });
        Check(result.Added == 1 && result.GeneratedIds == 1, "GUID duplicates must receive new IDs.");
        Check((Guid)table.Rows[1]["UID"] != old, "New GUID must differ from source.");
    }

    private static void TestCellCopyAndBroadcast()
    {
        DataTable table = Example();
        var source = GridCellEditor.CopyValues(new[]
        {
            new GridCellEditor.CellValue(0, 0, table.Rows[0]["BeltUID"])
        });
        Check(source == "0", "Copy must include only the selected cell, not the complete record.");
        var clipboard = GridCellEditor.ParseValues(source);
        Check(clipboard.Length == 1 && clipboard[0].Length == 1, "Single cell clipboard format.");
        var patches = Enumerable.Range(0, 2).Select(i =>
            new GridCellEditor.CellPatch(table.Rows[i], table.Columns["BeltOrder"]!, "1235343")).ToArray();
        int changed = GridCellEditor.Apply(patches);
        Check(changed == 2, "One copied number should fill many selected cells.");
        Check((int)table.Rows[0]["BeltOrder"] == 1235343 &&
              (int)table.Rows[1]["BeltOrder"] == 1235343, "Both target cells should be updated.");
        Check(table.Rows.Count == 2, "Cell pasting must never create a new record.");
        Check((int)table.Rows[0]["UID"] == 13 &&
              (string)table.Rows[0]["Recordname"] == "Prairie Cup013", "Other columns unchanged.");
    }

    private static void TestMultiColumnPaste()
    {
        DataTable table = Example();
        string input = "7\t40\r\n8\t41";
        var copied = GridCellEditor.ParseValues(input);
        var patches = new List<GridCellEditor.CellPatch>();
        for (int r = 0; r < copied.Length; r++)
        {
            patches.Add(new GridCellEditor.CellPatch(table.Rows[r], table.Columns["BeltUID"]!, copied[r][0]));
            patches.Add(new GridCellEditor.CellPatch(table.Rows[r], table.Columns["BeltOrder"]!, copied[r][1]));
        }
        Check(GridCellEditor.Apply(patches) == 4, "Multi-column grid paste count.");
        Check((int)table.Rows[0]["BeltUID"] == 7 &&
              (int)table.Rows[1]["BeltOrder"] == 41, "Pasted cells should match their coordinates.");
        Check(table.Rows.Count == 2, "Pasting cell blocks must not insert records.");
    }

    private static void TestProtectedKeysAndInvalidValues()
    {
        DataTable table = Example();
        try
        {
            GridCellEditor.Apply(new[] { new GridCellEditor.CellPatch(table.Rows[0], table.Columns["UID"]!, "999") });
            throw new Exception("Editing the primary key should have been blocked.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("primary key")) { }

        try
        {
            GridCellEditor.Apply(new[]
            {
                new GridCellEditor.CellPatch(table.Rows[0], table.Columns["BeltOrder"]!, "55"),
                new GridCellEditor.CellPatch(table.Rows[1], table.Columns["BeltUID"]!, "not a number")
            });
            throw new Exception("Invalid numeric paste should have been blocked.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Cannot paste")) { }
        Check((int)table.Rows[0]["BeltOrder"] == 13,
            "A failed multi-cell paste must not partially modify the table.");
    }

    private static void TestTextClipboardRoundtrip()
    {
        string value = "Has \"quotes\", tabs\tand\nnew lines";
        string text = GridCellEditor.CopyValues(new[]
        {
            new GridCellEditor.CellValue(0, 0, value)
        });
        var parsed = GridCellEditor.ParseValues(text);
        Check(parsed.Length == 1 && parsed[0].Length == 1 && parsed[0][0] == value,
            "Copied text must roundtrip, even with embedded quotes, tabs and new lines.");
    }

    private static void TestDeleteSelectedRecords()
    {
        var table = Example();
        var first = table.Rows[0];
        var second = table.Rows[1];
        // Highlighting more than one cell in a row must not delete it twice.
        var selected = RecordDeletion.EligibleRows(new[] { first, first, second, second });
        Check(selected.Count == 2, "Selected cells must resolve to distinct records.");
        Check(table.Rows.Cast<DataRow>().All(r => r.RowState == DataRowState.Unchanged),
            "Selecting cells must not change the database.");
        Check(RecordDeletion.Stage(table, new[] { first, first }) == 1,
            "Exactly one selected record must be staged for deletion.");
        Check(first.RowState == DataRowState.Deleted, "Selected record must be marked deleted.");
        Check(second.RowState == DataRowState.Unchanged, "Unselected record must be untouched.");
        Check((int)second["BeltOrder"] == 14, "Delete must not clear or alter unrelated data.");
        Check(table.GetChanges(DataRowState.Deleted)?.Rows.Count == 1,
            "Deletion must remain pending until Save Changes.");
    }

    private static void TestDeletePendingInsertedRecord()
    {
        var table = Example();
        var existing = table.Rows[0];
        var newRow = table.Rows.Add(15, "Prairie Cup015", 0, 15);
        Check(newRow.RowState == DataRowState.Added, "New record should be pending.");
        Check(RecordDeletion.Stage(table, new[] { newRow }) == 1,
            "Pending new record must be removable.");
        Check(newRow.RowState == DataRowState.Detached,
            "Uncommitted insertion must be removed from the local table.");
        Check(existing.RowState == DataRowState.Unchanged,
            "Deleting newly added record must preserve saved records.");
    }

    private static void TestDeleteWithoutPrimaryKey()
    {
        var table = Example();
        table.PrimaryKey = Array.Empty<DataColumn>();
        var first = table.Rows[0];
        var second = table.Rows[1];
        Check(RecordDeletion.Stage(table, new[] { first, first }) == 1,
            "Delete must be available on keyless tables.");
        Check(first.RowState == DataRowState.Deleted && second.RowState == DataRowState.Unchanged,
            "Only selected rows should be staged, irrespective of the primary key.");
    }

    private static void TestKeylessDeletePredicateSafety()
    {
        var table = new DataTable("Keyless");
        var uid = table.Columns.Add("Victor", typeof(int));
        var shortText = table.Columns.Add("Recordname", typeof(string));
        shortText.MaxLength = 80;
        table.Columns.Add("LongNote", typeof(string));
        table.Columns.Add("Attachment", typeof(byte[]));
        var nullable = table.Columns.Add("Optional", typeof(int));
        nullable.AllowDBNull = true;
        table.Rows.Add(-1, "Angle001", "uncomparable text", new byte[] { 1 }, DBNull.Value);
        table.Rows.Add(-1, "Angle001", "different long text", new byte[] { 2 }, DBNull.Value);
        table.Rows.Add(-1, "Angle002", "other long text", new byte[] { 3 }, DBNull.Value);
        table.AcceptChanges();

        var first = table.Rows[0];
        var second = table.Rows[1];
        var third = table.Rows[2];
        var match = RecordDeletion.BuildKeylessMatch(first);
        Check(match.Where.Contains("[Victor] = ?") &&
              match.Where.Contains("[Recordname] = ?") &&
              match.Where.Contains("[Optional] IS NULL"),
            "Keyless matching should include numeric, short text and null fields.");
        Check(!match.Where.Contains("LongNote") && !match.Where.Contains("Attachment"),
            "Binary and long text should never be compared in Access WHERE clauses.");
        Check(match.Parameters.Count == 2,
            "Null predicates must not consume a positional OLE DB parameter.");
        Check(RecordDeletion.SameKeylessValues(first, second, match.ComparedColumns),
            "Rows identical across comparable fields must be treated as a duplicate group.");
        second["Recordname"] = "Angle002";
        Check(RecordDeletion.SameKeylessValues(first, second, match.ComparedColumns),
            "Unsaved edits must not change the original values used to identify saved records.");
        Check(!RecordDeletion.SameKeylessValues(first, third, match.ComparedColumns),
            "Distinct original short text must not be grouped with a different record.");
        var unsupported = new DataTable("OnlyLongText");
        unsupported.Columns.Add("Memo", typeof(string));
        var row = unsupported.Rows.Add("This long text cannot be used as an Access equality key.");
        unsupported.AcceptChanges();
        try
        {
            RecordDeletion.BuildKeylessMatch(row);
            throw new Exception("Unsafe no-key deletion should be rejected.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("safely compare")) { }
    }

    private static void TestExcelQueryExportAndImport()
    {
        string filename = Path.Combine(Path.GetTempPath(), "MDBStudio-Excel-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            var results = new DataTable("SQL results");
            results.Columns.Add("BeltUID", typeof(int));
            results.Columns.Add("Recordname", typeof(string));
            results.Columns.Add("Comment", typeof(string));
            results.Rows.Add(42, "RSPWHayate Cup001", "=1+1");
            results.Rows.Add(43, "000045", "Sample");
            ExcelTransfer.ExportQuery(results, filename, "Championship Query");
            Check(File.Exists(filename), "Query export should create an XLSX workbook.");

            var names = ExcelTransfer.Sheets(filename);
            Check(names.Count == 1 && names[0] == "Championship Query",
                "Exported sheet name must match the query label.");
            var imported = ExcelTransfer.ReadSheet(filename, names[0]);
            Check(imported.Rows.Count == 2 && imported.Columns.Count == 3,
                "Workbook import should preserve dimensions.");
            Check(Convert.ToString(imported.Rows[0]["BeltUID"]) == "42",
                "Excel numeric cells must be readable.");
            Check(Convert.ToString(imported.Rows[0]["Recordname"]) == "RSPWHayate Cup001",
                "Imported text IDs must preserve their padded numeric suffix.");
            Check(Convert.ToString(imported.Rows[1]["Recordname"]) == "000045",
                "Excel text values with leading zeros must remain text.");
            Check(Convert.ToString(imported.Rows[0]["Comment"]) == "=1+1",
                "Literal text starting with equals must not become a formula.");
            using var workbook = new ClosedXML.Excel.XLWorkbook(filename);
            Check(!workbook.Worksheet(1).Cell(2, 3).HasFormula,
                "Query export must never treat database text as executable Excel formulas.");
        }
        finally
        {
            try { File.Delete(filename); } catch { }
        }
    }

    private static void TestExcelFileRejections()
    {
        try
        {
            ExcelTransfer.Sheets("not-an-excel-file.xls");
            throw new Exception("Legacy XLS should have been rejected.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains(".xlsx")) { }

        var empty = new DataTable("Empty Query");
        try
        {
            ExcelTransfer.ExportQuery(empty, Path.Combine(Path.GetTempPath(), "empty.xlsx"));
            throw new Exception("Exporting a query without column definitions should be rejected.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("SELECT query")) { }
    }

    private static void TestCsvQuotedImportAndExport()
    {
        string filename = Path.Combine(Path.GetTempPath(),
            "MDBStudio-Csv-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            // Quoted comma, embedded quote, embedded newline, UTF-8 and leading zeros.
            File.WriteAllText(filename,
                "UID,Recordname,Comment\r\n" +
                "000012,\"RSPW, Cup001\",\"Said \"\"hello\"\"\"\r\n" +
                "000013,RSPWHayate Cup002,\"Line one\nLine two\"\r\n" +
                "000014,水星,Text\r\n", new System.Text.UTF8Encoding(true));

            var table = CsvTransfer.ReadFile(filename);
            Check(table.Rows.Count == 3 && table.Columns.Count == 3,
                "CSV should preserve the expected number of records and columns.");
            Check((string)table.Rows[0]["UID"] == "000012",
                "Import must preserve leading zeros in CSV identifiers.");
            Check((string)table.Rows[0]["Recordname"] == "RSPW, Cup001",
                "Import must correctly read commas within quotes.");
            Check((string)table.Rows[0]["Comment"] == "Said \"hello\"",
                "Import must unescape doubled quotes.");
            Check((string)table.Rows[1]["Comment"] == "Line one\nLine two",
                "Import must correctly read multiline quoted CSV fields.");
            Check((string)table.Rows[2]["Recordname"] == "水星",
                "Import must preserve UTF-8 Unicode.");

            string export = Path.Combine(Path.GetTempPath(),
                "MDBStudio-Csv-Export-" + Guid.NewGuid().ToString("N") + ".csv");
            try
            {
                CsvTransfer.ExportQuery(table, export);
                var roundtrip = CsvTransfer.ReadFile(export);
                Check(roundtrip.Rows.Count == table.Rows.Count, "Query export should preserve CSV row count.");
                Check((string)roundtrip.Rows[0]["UID"] == "000012" &&
                      (string)roundtrip.Rows[1]["Comment"] == "Line one\nLine two",
                    "Query export must round-trip quoted CSV data.");
            }
            finally { try { File.Delete(export); } catch { } }
        }
        finally { try { File.Delete(filename); } catch { } }
    }

    private static void TestCsvMalformedRows()
    {
        string file = Path.Combine(Path.GetTempPath(),
            "MDBStudio-Csv-Invalid-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(file, "A,B\r\none,two\r\nonlyone\r\n");
            try
            {
                CsvTransfer.ReadFile(file);
                throw new Exception("CSV import with missing columns should fail.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("fields; expected")) { }
        }
        finally { try { File.Delete(file); } catch { } }
    }

    private static void TestCsvFormulaEscaping()
    {
        string file = Path.Combine(Path.GetTempPath(),
            "MDBStudio-Csv-Formula-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            var data = new DataTable("Results");
            data.Columns.Add("Value", typeof(string));
            data.Columns.Add("Amount", typeof(int));
            data.Rows.Add("=HYPERLINK(\"https://example.com\", \"x\")", -7);
            CsvTransfer.ExportQuery(data, file);
            var imported = CsvTransfer.ReadFile(file);
            Check(((string)imported.Rows[0]["Value"]).StartsWith("'="),
                "CSV export must neutralize spreadsheet formula-like text.");
            Check((string)imported.Rows[0]["Amount"] == "-7",
                "Negative numeric values must remain intact.");
        }
        finally { try { File.Delete(file); } catch { } }
    }

    private static void TestAccessBooleanSpreadsheetValues()
    {
        // Generic Access Yes/No fields accept -1 as True.
        var flag = new DataTable("Workers").Columns.Add("Active", typeof(bool));
        foreach (var value in new[] { "-1", "1", "True", "YES", "on", " t " })
            Check((bool)SpreadsheetImportValue.ConvertToColumn(value, flag),
                "Yes/No True conversion should work for genuine Boolean fields.");
        foreach (var value in new[] { "0", "False", "no", "off" })
            Check(!(bool)SpreadsheetImportValue.ConvertToColumn(value, flag),
                "Yes/No False conversion should work for genuine Boolean fields.");

        // The Victor field is an integer winner ID. TEW uses -1 when a
        // history entry is an angle rather than a match.
        var victor = new DataTable("tblShowHistory").Columns.Add("Victor", typeof(int));
        Check((int)SpreadsheetImportValue.ConvertToColumn("-1", victor) == -1,
            "Angle Victor -1 must remain the integer -1.");
        Check((int)SpreadsheetImportValue.ConvertToColumn("1235343", victor) == 1235343,
            "Match Victor worker IDs must be preserved.");
        Check((int)SpreadsheetImportValue.ConvertToColumn("0", victor) == 0,
            "A numeric zero must not be coerced to False.");

        var shortVictor = new DataTable("tblAngles").Columns.Add("Victor", typeof(short));
        Check((short)SpreadsheetImportValue.ConvertToColumn("-1", shortVictor) == -1,
            "Angle sentinel must also survive a signed 16-bit integer column.");

        // A Boolean Victor column cannot represent the distinct numeric
        // no-winner sentinel. The importer must reject rather than corrupt.
        var mistakenBoolean = new DataTable("tblShowHistory").Columns.Add("Victor", typeof(bool));
        try
        {
            SpreadsheetImportValue.ConvertToColumn("-1", mistakenBoolean);
            throw new Exception("TEW Victor -1 was silently changed into Boolean True.");
        }
        catch (FormatException ex) when (ex.Message.Contains("no match winner") &&
                                         ex.Message.Contains("numeric")) { }
    }

    private static void TestAccessBooleanCsvImport()
    {
        string filename = Path.Combine(Path.GetTempPath(),
            "MDBStudio-Angle-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            File.WriteAllText(filename,
                "Type,Victor,Active\r\nAngle,-1,-1\r\nMatch,346,0\r\n");
            var source = CsvTransfer.ReadFile(filename);
            Check(source.Rows.Count == 2, "Both angle and match rows must be imported.");

            var schema = new DataTable("tblShowHistory");
            var victor = schema.Columns.Add("Victor", typeof(int));
            var active = schema.Columns.Add("Active", typeof(bool));
            Check((int)SpreadsheetImportValue.ConvertToColumn(
                    (string)source.Rows[0]["Victor"], victor) == -1,
                "Angle -1 winner sentinel must be preserved from a real CSV file.");
            Check((int)SpreadsheetImportValue.ConvertToColumn(
                    (string)source.Rows[1]["Victor"], victor) == 346,
                "Match winner ID must be preserved from CSV.");
            Check((bool)SpreadsheetImportValue.ConvertToColumn(
                    (string)source.Rows[0]["Active"], active),
                "Genuine Access Yes/No -1 must still import as True.");
            Check(!(bool)SpreadsheetImportValue.ConvertToColumn(
                    (string)source.Rows[1]["Active"], active),
                "Genuine Access Yes/No 0 must still import as False.");
        }
        finally { try { File.Delete(filename); } catch { } }
    }

    private static void TestAccessBooleanCellPaste()
    {
        var table = new DataTable("tblShowHistory");
        var id = table.Columns.Add("UID", typeof(int));
        var victor = table.Columns.Add("Victor", typeof(int));
        var flag = table.Columns.Add("Active", typeof(bool));
        table.PrimaryKey = new[] { id };
        table.Rows.Add(1, 346, false);
        table.AcceptChanges();
        Check(GridCellEditor.Apply(new[] {
            new GridCellEditor.CellPatch(table.Rows[0], victor, "-1")
        }) == 1, "Pasting -1 into numeric Victor should update the angle sentinel.");
        Check((int)table.Rows[0][victor] == -1,
            "Cell paste must not coerce angle Victor -1 into Boolean True.");
        Check(GridCellEditor.Apply(new[] {
            new GridCellEditor.CellPatch(table.Rows[0], flag, "-1")
        }) == 1, "Pasting -1 into a genuine Yes/No field must work.");
        Check((bool)table.Rows[0][flag], "Real Boolean fields must still import True.");
    }

    private static void TestNumericClipboardAndLegacyCsvEscapes()
    {
        var table = new DataTable("tblMatchHistory");
        var victor = table.Columns.Add("Victor", typeof(int));
        var order = table.Columns.Add("BeltOrder", typeof(short));

        // Past exports may add a leading apostrophe to negative text numbers.
        // Other files can use lookalike Unicode minus signs and invisible
        // copy/paste marks. All of these represent exactly the integer -1.
        foreach (string value in new[] {
            "-1", " -1 ", "'-1", "' -1 ", "\u22121", "\uFF0D1",
            "\u20111", "\u20131", "\uFE631", "\uFEFF-1",
            "\u200B-1", "\u2060-1", "\u00A0-1\u00A0",
            "'\u22121", "'\uFF0D1"
        })
        {
            Check((int)SpreadsheetImportValue.ConvertToColumn(value, victor) == -1,
                "Victor=-1 should be preserved after safe numeric cleanup: " + value);
            Check((short)SpreadsheetImportValue.ConvertToColumn(value, order) == -1,
                "Negative short should also parse after safe cleanup.");
        }
        Check((int)SpreadsheetImportValue.ConvertToColumn("1235343", victor) == 1235343,
            "Normal match winner identifiers must remain unchanged.");
        Check((int)SpreadsheetImportValue.ConvertToColumn("1,234", victor) == 1234,
            "Valid thousands-formatted integers must still work.");

        var text = table.Columns.Add("OriginalText", typeof(string));
        Check((string)SpreadsheetImportValue.ConvertToColumn("'-1", text) == "'-1",
            "Imported text must never be modified by numeric-only normalization.");

        foreach (string unsafeValue in new[] { "'=1+2", " -SUM(1)", "12A", "--1" })
        {
            try
            {
                SpreadsheetImportValue.ConvertToColumn(unsafeValue, victor);
                throw new Exception("Invalid numeric value was silently accepted: " + unsafeValue);
            }
            catch (FormatException ex) when (ex.Message.Contains("U+") &&
                                             ex.Message.Contains("Normalized value")) { }
        }
    }

    private static void TestSignedNumericCsvExportRoundtrip()
    {
        string file = Path.Combine(Path.GetTempPath(),
            "MDBStudio-NumericCsv-" + Guid.NewGuid().ToString("N") + ".csv");
        try
        {
            var table = new DataTable("Old Save");
            table.Columns.Add("Victor", typeof(string));
            table.Columns.Add("Comment", typeof(string));
            table.Rows.Add("-1", "=HYPERLINK(\"https://example.com\",\"test\")");
            CsvTransfer.ExportQuery(table, file);
            var read = CsvTransfer.ReadFile(file);
            Check((string)read.Rows[0]["Victor"] == "-1",
                "CSV export must not add a quote to a safe signed number.");
            Check(((string)read.Rows[0]["Comment"]).StartsWith("'="),
                "Formula-like CSV text must still be protected.");
            var numeric = new DataTable("Destination").Columns.Add("Victor", typeof(int));
            Check((int)SpreadsheetImportValue.ConvertToColumn(
                (string)read.Rows[0]["Victor"], numeric) == -1,
                "A CSV file exported from MDB Studio must re-import numeric -1 intact.");

            // Even CSV files written by the old exporter with its leading
            // apostrophe should import safely into integer columns.
            File.WriteAllText(file, "Victor\r\n\"'-1\"\r\n");
            var oldFile = CsvTransfer.ReadFile(file);
            Check((int)SpreadsheetImportValue.ConvertToColumn(
                (string)oldFile.Rows[0]["Victor"], numeric) == -1,
                "Old MDB Studio CSV escaping must be accepted for numeric columns.");
        }
        finally { try { File.Delete(file); } catch { } }
    }

    private static void TestNumericImportDiagnostics()
    {
        var victor = new DataTable("Destination").Columns.Add("Victor", typeof(int));
        try
        {
            SpreadsheetImportValue.ConvertToColumn(" \u2212notnumber", victor);
            throw new Exception("Invalid numeric input must be rejected.");
        }
        catch (FormatException ex) when (
            ex.Message.Contains("U+2212") &&
            ex.Message.Contains("normalized", StringComparison.OrdinalIgnoreCase))
        {
            // Exact original Unicode minus U+2212 is included for diagnosis.
        }
    }

    private static void Check(bool condition, string failure)
    {
        if (!condition) throw new Exception(failure);
    }
}
