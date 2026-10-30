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

    private static void Check(bool condition, string failure)
    {
        if (!condition) throw new Exception(failure);
    }
}
