using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace MDBStudio;

/// <summary>
/// Converts an existing populated, keyed column only after testing it on an
/// isolated MDB/ACCDB file. Every existing row is validated before staging
/// and compared by primary key after ACE runs the DDL. A failed conversion
/// leaves the original file byte-for-byte unchanged.
/// </summary>
public sealed partial class AccessSession
{
    private void AlterFieldTypeSafely(
        string table, SchemaColumnInfo existing, AccessFieldDefinition target)
    {
        if (AccessFieldMetadata.IsSameType(existing, target))
            return;

        // Empty tables should use the complete table designer instead.
        // The DDL conversion path below is intended for populated tables.
        long rowCount = TableRecordCount(table);
        if (rowCount == 0)
            throw new InvalidOperationException(
                "This table has no records. Use Edit table design to change the field type. " +
                "MDB Studio can then rebuild the empty table safely, including its primary key.");

        var structure = InspectTable(table);
        var keys = structure.Columns.Where(c => c.PrimaryKey)
            .Select(c => c.Name).ToArray();
        if (keys.Length == 0)
            throw new InvalidOperationException(
                "Populated tables without a detected primary key cannot be verified " +
                "row by row after conversion. No changes were made.");

        ValidateStoredValues(table, existing.Name, target);

        string extension = Path.GetExtension(FilePath);
        string temporary = Path.Combine(Path.GetDirectoryName(FilePath)!,
            "MDBStudio-Convert-" + Guid.NewGuid().ToString("N") + extension);

        // A separate backup is retained for recovery even after successful
        // conversion. Staging file is always on the same volume as original.
        Backup();
        var before = new FileInfo(FilePath);
        long expectedLength = before.Length;
        DateTime expectedTimestamp = before.LastWriteTimeUtc;

        try
        {
            File.Copy(FilePath, temporary, overwrite: false);
            var staging = new AccessSession(temporary, _databasePassword);
            try
            {
                using var connection = staging.Open();
                using var command = new OleDbCommand(
                    AccessSchemaSql.AlterFieldType(table, target), connection);
                command.ExecuteNonQuery();
            }
            catch (OleDbException ex)
            {
                throw new InvalidOperationException(
                    "Microsoft Access rejected conversion of '" + existing.Name +
                    "' from " + existing.DataType + " to " + target.Kind +
                    " on a temporary database copy. No changes were applied to the original file. " +
                    "Even if every saved value can be represented, ACE may not support this " +
                    "specific direct ALTER COLUMN conversion. Try changing the field in " +
                    "Microsoft Access Design View, or add a separate new field and migrate " +
                    "the values before deleting the old one.\n\n" +
                    "Access engine detail: " + ex.Message, ex);
            }

            var changed = staging.InspectTable(table);
            var changedField = changed.Columns.FirstOrDefault(c =>
                c.Name.Equals(existing.Name, StringComparison.OrdinalIgnoreCase));
            if (changedField is null ||
                !AccessFieldMetadata.IsSameType(changedField, target))
                throw new InvalidOperationException(
                    "The Access engine returned a different field type or size than requested. " +
                    "The original database was not modified.");
            if (existing.Required != changedField.Required)
                throw new InvalidOperationException(
                    "The Access engine changed the field's Required property unexpectedly. " +
                    "The original database was not modified.");

            var afterKeys = changed.Columns.Where(c => c.PrimaryKey)
                .Select(c => c.Name).ToArray();
            if (!keys.SequenceEqual(afterKeys, StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The conversion changed the primary key. The original database was not modified.");

            if (!structure.IndexNames.SequenceEqual(changed.IndexNames,
                StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The conversion changed an existing index. The original database was not modified.");

            if (!ListTables().SequenceEqual(staging.ListTables(),
                StringComparer.OrdinalIgnoreCase) ||
                !ListSavedQueries().SequenceEqual(staging.ListSavedQueries(),
                    StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The conversion unexpectedly changed database objects. " +
                    "The original database was not modified.");

            CompareConvertedRows(staging, table, existing.Name, keys, target);

            // Another program may have changed the user's file while testing
            // the conversion. Never overwrite those newer changes.
            var latest = new FileInfo(FilePath);
            if (latest.Length != expectedLength ||
                latest.LastWriteTimeUtc != expectedTimestamp)
                throw new InvalidOperationException(
                    "The original database changed while testing the field conversion. " +
                    "No changes were applied. Close other editors and try again.");

            File.Replace(temporary, FilePath, destinationBackupFileName: null);
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

    private void ValidateStoredValues(
        string table, string column, AccessFieldDefinition target)
    {
        using var conn = Open();
        using var command = new OleDbCommand(
            "SELECT " + AccessSchemaSql.Name(column) +
            " FROM " + AccessSchemaSql.Name(table), conn);
        using var reader = command.ExecuteReader()
            ?? throw new InvalidOperationException("Could not inspect saved column values.");

        long row = 0;
        long badCount = 0;
        var examples = new List<string>();
        while (reader.Read())
        {
            row++;
            object value = reader.IsDBNull(0) ? DBNull.Value : reader.GetValue(0);
            try { _ = AccessFieldConversion.Normalize(value, target); }
            catch (Exception ex) when (ex is FormatException or InvalidOperationException)
            {
                badCount++;
                if (examples.Count < 5)
                    examples.Add("Record " + row.ToString("N0", CultureInfo.InvariantCulture) +
                        ": " + AccessFieldConversion.DisplaySample(value) +
                        ". " + ex.Message);
            }
        }

        if (badCount > 0)
            throw new InvalidOperationException(
                "Cannot convert '" + column + "' to " + target.Kind + ". " +
                badCount.ToString("N0", CultureInfo.InvariantCulture) +
                " saved records contain values that would not convert without loss. " +
                "No database changes were made.\n\n" +
                string.Join("\n", examples) +
                "\n\nCorrect these values or choose a compatible field type before retrying.");
    }

    private void CompareConvertedRows(
        AccessSession staged, string table, string column, IReadOnlyList<string> keys,
        AccessFieldDefinition target)
    {
        // The primary key remains unchanged, so both readers can return an
        // identically ordered full table without holding all records in memory.
        string keyProjection = string.Join(", ", keys.Select(AccessSchemaSql.Name));
        string field = AccessSchemaSql.Name(column);
        string sql = "SELECT " + keyProjection + ", " + field +
            " FROM " + AccessSchemaSql.Name(table) +
            " ORDER BY " + keyProjection;

        using var beforeConn = Open();
        using var afterConn = staged.Open();
        using var beforeCmd = new OleDbCommand(sql, beforeConn);
        using var afterCmd = new OleDbCommand(sql, afterConn);
        using var original = beforeCmd.ExecuteReader()
            ?? throw new InvalidOperationException("Could not verify original records.");
        using var converted = afterCmd.ExecuteReader()
            ?? throw new InvalidOperationException("Could not verify converted records.");

        long row = 0;
        while (true)
        {
            bool hasBefore = original.Read();
            bool hasAfter = converted.Read();
            if (hasBefore != hasAfter)
                throw new InvalidOperationException(
                    "The conversion changed the number of records. The original database was not modified.");
            if (!hasBefore) break;
            row++;

            for (int i = 0; i < keys.Count; i++)
            {
                if (!Equals(original.GetValue(i), converted.GetValue(i)))
                    throw new InvalidOperationException(
                        "The conversion changed a record's primary key at row " +
                        row.ToString("N0", CultureInfo.InvariantCulture) +
                        ". The original database was not modified.");
            }

            object left = original.GetValue(keys.Count);
            object right = converted.GetValue(keys.Count);
            if (!AccessFieldConversion.Matches(left, right, target))
                throw new InvalidOperationException(
                    "Access converted a value differently than expected in '" + column +
                    "' at row " + row.ToString("N0", CultureInfo.InvariantCulture) +
                    ". Original: " + AccessFieldConversion.DisplaySample(left) +
                    ", converted: " + AccessFieldConversion.DisplaySample(right) +
                    ". No changes were made to the original database.");
        }
    }
}
