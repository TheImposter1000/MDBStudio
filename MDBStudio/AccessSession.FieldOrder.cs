using System.Data;
using System.IO;
using System.Runtime.InteropServices;

namespace MDBStudio;

/// <summary>
/// Changes actual Access TableDef.Fields.OrdinalPosition, including on
/// populated tables. DAO supports editing this metadata directly, so there is
/// no reason to rebuild a table, drop a field, or modify any records.
/// </summary>
public sealed partial class AccessSession
{
    public IReadOnlyList<string> ReadDesignFieldOrder(string table)
    {
        AccessSchemaSql.Name(table);
        var structure = InspectTable(table);
        try
        {
            return AccessFieldOrder.Validate(
                structure.Columns.Select(c => c.Name).ToArray(),
                ReadDaoFieldOrder(FilePath, table, _databasePassword));
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException
            or MissingMethodException or TypeLoadException)
        {
            // The ordinary schema viewer still works when DAO is not installed.
            // Saving a physical reorder will report that DAO is required.
            return structure.Columns.Select(c => c.Name).ToArray();
        }
    }

    public void ReorderTableFields(string table, IReadOnlyList<string> requested)
    {
        AccessSchemaSql.Name(table);
        var structure = InspectTable(table);
        var existingNames = structure.Columns.Select(c => c.Name).ToArray();
        var wanted = AccessFieldOrder.Validate(existingNames, requested);
        var current = ReadDaoFieldOrder(FilePath, table, _databasePassword);
        AccessFieldOrder.Validate(current, wanted);

        if (AccessFieldOrder.Same(current, wanted)) return;

        string temporary = Path.Combine(Path.GetDirectoryName(FilePath)!,
            "MDBStudio-FieldOrder-" + Guid.NewGuid().ToString("N") +
            Path.GetExtension(FilePath));

        // Silent backup, no modal confirmation. The original is changed
        // only after the temporary copy proves it can store the requested order.
        Backup();
        var source = new FileInfo(FilePath);
        long oldLength = source.Length;
        DateTime oldStamp = source.LastWriteTimeUtc;

        try
        {
            File.Copy(FilePath, temporary, overwrite: false);
            WriteDaoFieldOrder(temporary, table, wanted, _databasePassword);
            var actual = ReadDaoFieldOrder(temporary, table, _databasePassword);
            if (!AccessFieldOrder.Same(actual, wanted))
                throw new InvalidOperationException(
                    "The Access engine did not retain the requested field order. " +
                    "The original database was not changed.");

            // Changing OrdinalPosition must not change fields, keys, records,
            // or other objects. Verify the staging file before replacement.
            var staged = new AccessSession(temporary, _databasePassword);
            var stagedSchema = staged.InspectTable(table);
            var beforeByName = structure.Columns.ToDictionary(c => c.Name,
                StringComparer.OrdinalIgnoreCase);
            if (stagedSchema.Columns.Count != structure.Columns.Count ||
                stagedSchema.Columns.Any(c =>
                    !beforeByName.TryGetValue(c.Name, out var prior) ||
                    c.DataType != prior.DataType || c.Length != prior.Length ||
                    c.Required != prior.Required || c.PrimaryKey != prior.PrimaryKey))
                throw new InvalidOperationException(
                    "The Access engine unexpectedly altered a field definition. " +
                    "The original database was not changed.");

            if (!ListTables().SequenceEqual(staged.ListTables(),
                StringComparer.OrdinalIgnoreCase) ||
                !ListSavedQueries().SequenceEqual(staged.ListSavedQueries(),
                    StringComparer.OrdinalIgnoreCase) ||
                TableRecordCount(table) != staged.TableRecordCount(table))
                throw new InvalidOperationException(
                    "The staging database did not retain its existing objects or record count. " +
                    "The original database was not changed.");

            source.Refresh();
            if (source.Length != oldLength || source.LastWriteTimeUtc != oldStamp)
                throw new InvalidOperationException(
                    "The database changed while the new field order was being prepared. " +
                    "No changes were applied.");

            // Atomic replacement on the same volume.
            File.Replace(temporary, FilePath, destinationBackupFileName: null);
        }
        catch (COMException ex)
        {
            throw new InvalidOperationException(
                "Microsoft Access DAO could not change the field order. " +
                "The existing database has not been modified. " +
                "Check the 64-bit Microsoft Access Database Engine installation " +
                "and ensure the database is not open exclusively in another application. " +
                "Detail: " + ex.Message, ex);
        }
        finally
        {
            if (File.Exists(temporary))
            {
                try { File.Delete(temporary); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    private static IReadOnlyList<string> ReadDaoFieldOrder(
        string file, string table, string? password)
    {
        return WithDaoFields(file, table, password, fields =>
        {
            var positions = new List<(string Name, int Ordinal)>();
            int count = Convert.ToInt32(fields.Count);
            for (int i = 0; i < count; i++)
            {
                object? field = null;
                try
                {
                    field = fields[i];
                    dynamic item = field!;
                    positions.Add(((string)item.Name, (int)item.OrdinalPosition));
                }
                finally { ReleaseDao(field); }
            }
            return positions.OrderBy(p => p.Ordinal)
                .ThenBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
                .Select(p => p.Name).ToArray();
        });
    }

    private static void WriteDaoFieldOrder(
        string file, string table, IReadOnlyList<string> ordered, string? password)
    {
        _ = WithDaoFields(file, table, password, fields =>
        {
            for (int i = 0; i < ordered.Count; i++)
            {
                object? field = null;
                try
                {
                    field = fields[ordered[i]];
                    dynamic item = field!;
                    item.OrdinalPosition = i;
                }
                finally { ReleaseDao(field); }
            }
            fields.Refresh();
            return true;
        });
    }

    private static T WithDaoFields<T>(
        string path, string table, string? password,
        Func<dynamic, T> operation)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("DAO field ordering requires Windows.");

        Type engineType = Type.GetTypeFromProgID("DAO.DBEngine.120")
            ?? throw new InvalidOperationException(
                "The 64-bit Microsoft Access DAO component is not installed. " +
                "Install or repair the matching Microsoft Access Database Engine " +
                "to change the actual field order.");
        object engine = Activator.CreateInstance(engineType)
            ?? throw new InvalidOperationException("Could not start Microsoft Access DAO.");
        object? database = null;
        object? tableDefinition = null;
        object? fieldCollection = null;
        try
        {
            dynamic dao = engine;
            // The fourth OpenDatabase argument is the Access connection
            // string; do not write passwords to settings, errors or logs.
            database = dao.OpenDatabase(path, false, false,
                string.IsNullOrEmpty(password) ? "" : ";PWD=" + password);
            dynamic db = database;
            tableDefinition = db.TableDefs[table];
            dynamic definition = tableDefinition;
            fieldCollection = definition.Fields;
            return operation(fieldCollection);
        }
        finally
        {
            try
            {
                ReleaseDao(fieldCollection);
                ReleaseDao(tableDefinition);
                if (database is not null)
                {
                    try { ((dynamic)database).Close(); }
                    finally { ReleaseDao(database); }
                }
            }
            finally { ReleaseDao(engine); }
        }
    }

    private static void ReleaseDao(object? instance)
    {
        if (instance is not null && Marshal.IsComObject(instance))
            Marshal.FinalReleaseComObject(instance);
    }
}
