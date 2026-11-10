using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Globalization;

namespace MDBStudio;

/// <summary>
/// Fully replaces the design of an EMPTY Access table. DDL is performed on a
/// separate copy of the database and the original is atomically replaced only
/// when all operations and validation have succeeded. Populated tables are
/// never rebuilt, and an untouched restorable original backup is retained.
/// </summary>
public sealed partial class AccessSession
{
    public long TableRecordCount(string name)
    {
        AccessSchemaSql.Name(name);
        if (!ListTables().Contains(name, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The selected table no longer exists.");
        using var conn = Open();
        using var count = new OleDbCommand(
            "SELECT COUNT(*) FROM " + AccessSchemaSql.Name(name), conn);
        return Convert.ToInt64(count.ExecuteScalar(), CultureInfo.InvariantCulture);
    }

    public EmptyTableRedesignPlan PlanEmptyTableRedesign(
        string name, IReadOnlyList<AccessFieldDefinition> fields)
    {
        AccessSchemaSql.Name(name);
        if (!ListTables().Contains(name, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException("The table no longer exists.");

        using var conn = Open();
        EnsureNoRelationships(conn, name);
        using var count = new OleDbCommand(
            "SELECT COUNT(*) FROM " + AccessSchemaSql.Name(name), conn);
        long rows = Convert.ToInt64(count.ExecuteScalar(), CultureInfo.InvariantCulture);
        // Read the saved schema, not the currently visible page.
        var indexes = ReadSecondaryIndexes(conn, name);
        return EmptyTableRedesign.Plan(name, rows, fields, indexes);
    }

    public string RedesignEmptyTable(EmptyTableRedesignPlan proposed)
    {
        string name = proposed.TableName;
        // Do not trust stale UI state. Revalidate the table and proposed
        // schema against the actual file immediately before making a copy.
        EmptyTableRedesignPlan current = PlanEmptyTableRedesign(name, proposed.Fields);
        if (!SameIndexes(current.RetainedIndexes, proposed.RetainedIndexes) ||
            !current.RemovedIndexNames.SequenceEqual(proposed.RemovedIndexNames,
                StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The table's indexes have changed since you opened the designer. " +
                "Reload the table and review the new design before trying again.");

        string folder = Path.GetDirectoryName(FilePath)!;
        string temporary = Path.Combine(folder,
            "MDBStudio-Redesign-" + Guid.NewGuid().ToString("N") +
            Path.GetExtension(FilePath));

        // A full backup is made before the original could be replaced.
        // The backup remains available even when a migration cannot finish.
        string backup = Backup();
        FileInfo originalInfo = new(FilePath);
        long expectedLength = originalInfo.Length;
        DateTime expectedStamp = originalInfo.LastWriteTimeUtc;

        try
        {
            File.Copy(FilePath, temporary, overwrite: false);

            // All DROP TABLE / CREATE TABLE commands run on the copied file.
            // A failure, a relationship constraint, or a locked file leaves
            // the user's original database untouched.
            var staging = new AccessSession(temporary, _databasePassword);
            using (var stagedConnection = staging.Open())
            {
                using (var drop = new OleDbCommand(
                    AccessSchemaSql.DeleteTable(name), stagedConnection))
                    drop.ExecuteNonQuery();

                using (var create = new OleDbCommand(
                    current.CreateTableSql, stagedConnection))
                    create.ExecuteNonQuery();

                foreach (var index in current.RetainedIndexes)
                {
                    string createIndex = AccessSchemaSql.CreateIndex(
                        name, index.Name, index.Columns, index.Unique);
                    using var command = new OleDbCommand(createIndex, stagedConnection);
                    command.ExecuteNonQuery();
                }

                using var verifyCount = new OleDbCommand(
                    "SELECT COUNT(*) FROM " + AccessSchemaSql.Name(name),
                    stagedConnection);
                long newRows = Convert.ToInt64(verifyCount.ExecuteScalar(),
                    CultureInfo.InvariantCulture);
                if (newRows != 0)
                    throw new InvalidOperationException(
                        "Verification failed: the redesigned table is not empty.");
            }

            var stagedSchema = staging.InspectTable(name);
            if (stagedSchema.Columns.Count != current.Fields.Count ||
                !stagedSchema.Columns.Select(c => c.Name).SequenceEqual(
                    current.Fields.Select(f => f.Name),
                    StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Verification failed: the replacement table does not match the proposed fields.");

            var expectedKeys = current.Fields.Where(f => f.PrimaryKey)
                .Select(f => f.Name).ToArray();
            var actualKeys = stagedSchema.Columns.Where(c => c.PrimaryKey)
                .Select(c => c.Name).ToArray();
            if (!expectedKeys.SequenceEqual(actualKeys,
                StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Verification failed: the replacement table does not contain the expected primary key.");

            // Reject a copied-file rewrite if it accidentally removes other
            // objects. If this happens it is safe: only the staging copy changed.
            var expectedQueries = ListSavedQueries();
            var stagedQueries = staging.ListSavedQueries();
            if (!expectedQueries.SequenceEqual(stagedQueries,
                StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Verification failed: saved query objects changed on the temporary copy.");

            var expectedTables = ListTables();
            var stagedTables = staging.ListTables();
            if (!expectedTables.SequenceEqual(stagedTables,
                StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Verification failed: the database table list changed unexpectedly.");

            // A separate process could have edited the source during staging.
            // Do not overwrite that process's records or structure.
            FileInfo latest = new(FilePath);
            if (latest.Length != expectedLength ||
                latest.LastWriteTimeUtc != expectedStamp)
                throw new InvalidOperationException(
                    "The original database changed while the redesign was being prepared. " +
                    "No changes were applied; close other database applications and try again.");

            // Same-volume File.Replace supplies atomic replacement semantics.
            // Backup() above leaves a separate original file in MDBStudio Backups.
            File.Replace(temporary, FilePath, destinationBackupFileName: null);
            return backup;
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

    private static void EnsureNoRelationships(OleDbConnection connection, string table)
    {
        try
        {
            var foreignKeys = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null);
            if (foreignKeys is null)
                throw new InvalidOperationException(
                    "The Access provider did not return relationship metadata. " +
                    "Complete table redesign was cancelled to avoid losing constraints.");
            foreach (DataRow row in foreignKeys.Rows)
            {
                string parent = foreignKeys.Columns.Contains("PK_TABLE_NAME")
                    ? Convert.ToString(row["PK_TABLE_NAME"], CultureInfo.InvariantCulture) ?? ""
                    : "";
                string child = foreignKeys.Columns.Contains("FK_TABLE_NAME")
                    ? Convert.ToString(row["FK_TABLE_NAME"], CultureInfo.InvariantCulture) ?? ""
                    : "";
                if (parent.Equals(table, StringComparison.OrdinalIgnoreCase) ||
                    child.Equals(table, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException(
                        "This table participates in a relationship. Redesign it using " +
                        "supported individual field operations instead. MDB Studio will " +
                        "not rebuild a linked table or remove a relationship automatically.");
            }
        }
        catch (OleDbException ex)
        {
            throw new InvalidOperationException(
                "Could not verify table relationships. No redesign has been applied.", ex);
        }
    }

    private static bool SameIndexes(
        IReadOnlyList<AccessRedesignIndex> x, IReadOnlyList<AccessRedesignIndex> y)
    {
        if (x.Count != y.Count) return false;
        return x.Zip(y).All(pair =>
            string.Equals(pair.First.Name, pair.Second.Name,
                StringComparison.OrdinalIgnoreCase) &&
            pair.First.Unique == pair.Second.Unique &&
            pair.First.Columns.SequenceEqual(pair.Second.Columns,
                StringComparer.OrdinalIgnoreCase));
    }

    private static IReadOnlyList<AccessRedesignIndex> ReadSecondaryIndexes(
        OleDbConnection connection, string table)
    {
        var schema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Indexes, null);
        if (schema is null)
            throw new InvalidOperationException(
                "Access did not return index metadata. Redesign was cancelled to protect existing indexes.");
        var required = new[] { "INDEX_NAME", "TABLE_NAME", "COLUMN_NAME", "PRIMARY_KEY", "UNIQUE" };
        if (required.Any(column => !schema.Columns.Contains(column)))
            throw new InvalidOperationException(
                "Access did not return complete index metadata. Redesign was cancelled.");

        var values = schema.Rows.Cast<DataRow>()
            .Where(row => string.Equals(
                Convert.ToString(row["TABLE_NAME"], CultureInfo.InvariantCulture),
                table, StringComparison.OrdinalIgnoreCase))
            .Where(row => row["PRIMARY_KEY"] == DBNull.Value ||
                !Convert.ToBoolean(row["PRIMARY_KEY"], CultureInfo.InvariantCulture))
            .ToArray();
        var indexes = new List<AccessRedesignIndex>();
        foreach (var group in values.GroupBy(row =>
            Convert.ToString(row["INDEX_NAME"], CultureInfo.InvariantCulture) ?? "",
            StringComparer.OrdinalIgnoreCase))
        {
            if (string.IsNullOrWhiteSpace(group.Key))
                throw new InvalidOperationException(
                    "An Access index has no readable name. Redesign was cancelled.");
            var rows = group.OrderBy(row =>
                schema.Columns.Contains("ORDINAL_POSITION") &&
                row["ORDINAL_POSITION"] != DBNull.Value
                    ? Convert.ToInt32(row["ORDINAL_POSITION"], CultureInfo.InvariantCulture)
                    : 0).ToArray();
            var columns = rows.Select(row =>
                Convert.ToString(row["COLUMN_NAME"], CultureInfo.InvariantCulture) ?? "")
                .ToArray();
            if (columns.Any(string.IsNullOrWhiteSpace))
                throw new InvalidOperationException(
                    "Index " + group.Key + " has unreadable fields; redesign was cancelled.");
            bool unique = rows[0]["UNIQUE"] != DBNull.Value &&
                Convert.ToBoolean(rows[0]["UNIQUE"], CultureInfo.InvariantCulture);
            indexes.Add(new AccessRedesignIndex(group.Key, columns, unique));
        }
        return indexes.OrderBy(i => i.Name, StringComparer.OrdinalIgnoreCase).ToArray();
    }
}
