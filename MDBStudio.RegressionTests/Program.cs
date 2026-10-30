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
        TestDeleteRequiresPrimaryKey();
        Console.WriteLine("Paste identity regression tests passed.");
        return 0;
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

    private static void TestDeleteRequiresPrimaryKey()
    {
        var table = Example();
        table.PrimaryKey = Array.Empty<DataColumn>();
        try
        {
            RecordDeletion.Stage(table, new[] { table.Rows[0] });
            throw new Exception("Expected delete without a primary key to be rejected.");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("primary key")) { }
        Check(table.Rows[0].RowState == DataRowState.Unchanged,
            "Unsupported deletion must leave all records unchanged.");
    }

    private static void Check(bool condition, string failure)
    {
        if (!condition) throw new Exception(failure);
    }
}
