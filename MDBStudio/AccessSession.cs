using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Text.RegularExpressions;

namespace MDBStudio;

public sealed class DatabasePasswordRequiredException : InvalidOperationException
{
    public DatabasePasswordRequiredException(bool suppliedPassword)
        : base(suppliedPassword
            ? "The supplied database password was not accepted, or this file uses a format that Microsoft ACE cannot read."
            : "This database may require a password. Enter a password only if you already know it. Some unsupported or damaged files may produce the same error.")
    {
    }
}

/// <summary>
/// Local ACE OLE DB access. Identifiers are validated, record values are parameterized,
/// and writes require a primary key and a backup of the original database.
/// </summary>
public sealed class AccessSession
{
    private readonly string _provider;
    private readonly string? _databasePassword;
    private bool _hasBackup;
    public const int PreviewRowLimit = 2000;
    public string FilePath { get; }

    public AccessSession(string path, string? databasePassword = null)
    {
        _databasePassword = databasePassword;
        FilePath = Path.GetFullPath(path);
        if (!File.Exists(FilePath))
            throw new FileNotFoundException("Database file not found.", FilePath);

        string extension = Path.GetExtension(FilePath);
        if (!extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".accdb", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Choose an MDB or ACCDB file.");

        var errors = new List<string>();
        bool passwordRejected = false;
        foreach (var provider in new[] { "Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0" })
        {
            try
            {
                using var conn = NewConnection(provider);
                conn.Open();
                _provider = provider;
                return;
            }
            catch (Exception ex) when (ex is OleDbException or InvalidOperationException)
            {
                if (ex.Message.Contains("Not a valid password", StringComparison.OrdinalIgnoreCase))
                    passwordRejected = true;
                errors.Add(provider + ": " + ex.Message);
            }
        }
        if (passwordRejected)
            throw new DatabasePasswordRequiredException(!string.IsNullOrEmpty(databasePassword));

        throw new InvalidOperationException(
            "Could not open the database with either 64-bit Microsoft ACE OLE DB provider.\n" +
            "Check that the file exists, is an Access database, and is not open exclusively.\n\n" +
            string.Join("\n\n", errors));
    }

    private OleDbConnection NewConnection(string provider)
    {
        var builder = new OleDbConnectionStringBuilder
        {
            Provider = provider,
            DataSource = FilePath,
            PersistSecurityInfo = false
        };
        if (!string.IsNullOrEmpty(_databasePassword))
            builder["Jet OLEDB:Database Password"] = _databasePassword;
        return new OleDbConnection(builder.ConnectionString);
    }

    private OleDbConnection Open()
    {
        var conn = NewConnection(_provider);
        try
        {
            conn.Open();
            return conn;
        }
        catch
        {
            conn.Dispose();
            throw;
        }
    }

    private static string QuoteName(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Contains(']') || name.Contains('\0'))
            throw new ArgumentException("Unsupported Access object name.");
        return "[" + name + "]";
    }

    public IReadOnlyList<string> ListTables()
    {
        using var conn = Open();
        var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        if (schema is null) return Array.Empty<string>();

        return schema.Rows.Cast<DataRow>()
            .Where(row => string.Equals(Convert.ToString(row["TABLE_TYPE"]), "TABLE", StringComparison.OrdinalIgnoreCase))
            .Select(row => Convert.ToString(row["TABLE_NAME"]) ?? "")
            .Where(name => !name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) &&
                           !name.StartsWith("~", StringComparison.Ordinal) &&
                           !string.IsNullOrWhiteSpace(name))
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    public DataTable Load(string name)
    {
        using var conn = Open();
        using var adapter = new OleDbDataAdapter(
            "SELECT TOP " + PreviewRowLimit + " * FROM " + QuoteName(name), conn)
        {
            MissingSchemaAction = MissingSchemaAction.AddWithKey
        };
        var table = new DataTable(name) { Locale = CultureInfo.InvariantCulture };
        adapter.Fill(table);

        if (table.PrimaryKey.Length == 0)
        {
            try
            {
                var primaryKeys = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys, new object?[] { null, null, name });
                if (primaryKeys is not null)
                {
                    var keys = primaryKeys.Rows.Cast<DataRow>()
                        .OrderBy(r => r.Table.Columns.Contains("ORDINAL") && r["ORDINAL"] != DBNull.Value
                            ? Convert.ToInt32(r["ORDINAL"], CultureInfo.InvariantCulture) : 0)
                        .Select(r => Convert.ToString(r["COLUMN_NAME"]) ?? "")
                        .Where(column => table.Columns.Contains(column))
                        .Select(column => table.Columns[column]!)
                        .ToArray();
                    if (keys.Length > 0) table.PrimaryKey = keys;
                }
            }
            catch (OleDbException)
            {
                // Without an identifiable key the UI intentionally stays read-only.
            }
        }
        return table;
    }

    private static OleDbType GetOleDbType(Type type, int maxLength)
    {
        if (type == typeof(string))
            return maxLength < 0 || maxLength > 255 ? OleDbType.LongVarWChar : OleDbType.VarWChar;
        if (type == typeof(int)) return OleDbType.Integer;
        if (type == typeof(short)) return OleDbType.SmallInt;
        if (type == typeof(long)) return OleDbType.BigInt;
        if (type == typeof(byte)) return OleDbType.UnsignedTinyInt;
        if (type == typeof(bool)) return OleDbType.Boolean;
        if (type == typeof(DateTime)) return OleDbType.DBTimeStamp;
        if (type == typeof(decimal)) return OleDbType.Decimal;
        if (type == typeof(float)) return OleDbType.Single;
        if (type == typeof(double)) return OleDbType.Double;
        if (type == typeof(Guid)) return OleDbType.Guid;
        if (type == typeof(byte[])) return OleDbType.LongVarBinary;
        return OleDbType.Variant;
    }

    private static void AddParameter(OleDbCommand command, DataColumn column, object value)
    {
        var parameter = new OleDbParameter
        {
            OleDbType = GetOleDbType(column.DataType, column.MaxLength),
            Value = value == DBNull.Value ? DBNull.Value : value
        };
        if (parameter.OleDbType == OleDbType.VarWChar && column.MaxLength > 0)
            parameter.Size = column.MaxLength;
        command.Parameters.Add(parameter);
    }

    public string Backup()
    {
        var folder = Path.Combine(Path.GetDirectoryName(FilePath)!, "MDBStudio Backups");
        Directory.CreateDirectory(folder);
        string destination = Path.Combine(folder,
            Path.GetFileNameWithoutExtension(FilePath) +
            "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fffffff", CultureInfo.InvariantCulture) +
            Path.GetExtension(FilePath));
        File.Copy(FilePath, destination, overwrite: false);
        _hasBackup = true;
        return destination;
    }

    public int Save(DataTable table)
    {
        var modified = table.Rows.Cast<DataRow>().Where(row =>
            row.RowState is DataRowState.Added or DataRowState.Modified or DataRowState.Deleted).ToArray();
        if (modified.Length == 0) return 0;

        if (table.PrimaryKey.Length == 0)
            throw new InvalidOperationException("Editing is disabled for tables without a detectable primary key.");

        foreach (var row in modified.Where(r => r.RowState == DataRowState.Modified))
        {
            if (table.PrimaryKey.Any(key =>
                !Equals(row[key, DataRowVersion.Original], row[key, DataRowVersion.Current])))
                throw new InvalidOperationException("Editing primary key values is not supported in this preview.");
        }

        // Important: file copy is done before a database connection and before any mutation.
        if (!_hasBackup) Backup();
        using var conn = Open();
        using var transaction = conn.BeginTransaction();
        try
        {
            int count = 0;
            var writable = table.Columns.Cast<DataColumn>()
                .Where(c => !c.AutoIncrement && !c.ReadOnly && string.IsNullOrEmpty(c.Expression))
                .ToArray();

            foreach (var row in modified.OrderBy(r =>
                r.RowState == DataRowState.Deleted ? 0 :
                r.RowState == DataRowState.Modified ? 1 : 2))
            {
                using var command = new OleDbCommand { Connection = conn, Transaction = transaction };
                DataColumn[] columns;
                if (row.RowState == DataRowState.Deleted)
                {
                    command.CommandText = "DELETE FROM " + QuoteName(table.TableName) + " WHERE " +
                        string.Join(" AND ", table.PrimaryKey.Select(k => QuoteName(k.ColumnName) + " = ?"));
                    foreach (var key in table.PrimaryKey)
                        AddParameter(command, key, row[key, DataRowVersion.Original]);
                }
                else if (row.RowState == DataRowState.Modified)
                {
                    columns = writable.Where(c => !table.PrimaryKey.Contains(c) &&
                        !Equals(row[c, DataRowVersion.Original], row[c, DataRowVersion.Current])).ToArray();
                    if (columns.Length == 0) continue;
                    command.CommandText = "UPDATE " + QuoteName(table.TableName) + " SET " +
                        string.Join(", ", columns.Select(c => QuoteName(c.ColumnName) + " = ?")) + " WHERE " +
                        string.Join(" AND ", table.PrimaryKey.Select(k => QuoteName(k.ColumnName) + " = ?"));
                    foreach (var col in columns) AddParameter(command, col, row[col, DataRowVersion.Current]);
                    foreach (var key in table.PrimaryKey) AddParameter(command, key, row[key, DataRowVersion.Original]);
                }
                else
                {
                    columns = writable;
                    if (columns.Length == 0)
                        throw new InvalidOperationException("This table has no writable columns.");
                    command.CommandText = "INSERT INTO " + QuoteName(table.TableName) + " (" +
                        string.Join(", ", columns.Select(c => QuoteName(c.ColumnName))) + ") VALUES (" +
                        string.Join(", ", columns.Select(_ => "?")) + ")";
                    foreach (var col in columns) AddParameter(command, col, row[col, DataRowVersion.Current]);
                }

                int affected = command.ExecuteNonQuery();
                if (row.RowState != DataRowState.Added && affected != 1)
                    throw new DBConcurrencyException("Expected exactly one changed record, but got " + affected + ".");
                count += affected;
            }
            transaction.Commit();
            return count;
        }
        catch
        {
            try { transaction.Rollback(); } catch { /* preserve original exception */ }
            throw;
        }
    }

    public DataTable Select(string sql)
    {
        if (string.IsNullOrWhiteSpace(sql) ||
            !Regex.IsMatch(sql, @"^\s*SELECT\s", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
            Regex.IsMatch(sql, @"\b(INTO|UPDATE|DELETE|INSERT|EXEC|DROP|ALTER|CREATE)\b",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant) ||
            sql.Contains(';'))
            throw new InvalidOperationException("Only a single read-only SELECT without INTO is permitted in this preview.");

        using var conn = Open();
        using var adapter = new OleDbDataAdapter(sql, conn);
        var result = new DataTable("SQL results");
        adapter.Fill(0, PreviewRowLimit, result);
        return result;
    }
}
