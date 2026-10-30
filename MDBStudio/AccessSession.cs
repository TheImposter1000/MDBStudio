using System.IO;
using System.Text;
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
    public const int DefaultPageSize = 5000;
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

    /// <summary>
    /// Fetches one page of a table, or every row if pageSize is zero.
    /// No fixed record ceiling: paging prevents large MDBs from exhausting UI memory.
    /// Pages are ordered by the table primary key where one is available.
    /// </summary>
    public TablePage LoadPage(string name, int pageIndex = 0, int pageSize = DefaultPageSize)
    {
        if (pageIndex < 0 || pageSize < 0)
            throw new ArgumentOutOfRangeException(nameof(pageIndex), "Page index and page size cannot be negative.");

        using var conn = Open();
        using var countCommand = new OleDbCommand("SELECT COUNT(*) FROM " + QuoteName(name), conn);
        long total = Convert.ToInt64(countCommand.ExecuteScalar(), CultureInfo.InvariantCulture);

        if (pageSize == 0) pageIndex = 0;
        else if (total > 0 && (long)pageIndex * pageSize >= total)
            pageIndex = (int)Math.Min(int.MaxValue, (total - 1) / pageSize);

        long start = (long)pageIndex * pageSize;
        if (start > int.MaxValue)
            throw new InvalidOperationException("This page is beyond the supported Access record offset.");

        // Access SQL has no OFFSET syntax. DbDataAdapter supports startRecord/maximum
        // and lets the UI page across the entire result set without a TOP 2000 ceiling.
        // When there is a primary key, ordering gives stable page boundaries.
        var primaryKeyNames = FindPrimaryKeys(conn, name);
        string orderBy = primaryKeyNames.Count == 0 ? "" : " ORDER BY " +
            string.Join(", ", primaryKeyNames.Select(QuoteName));
        string sql = "SELECT * FROM " + QuoteName(name) + orderBy;

        using var adapter = new OleDbDataAdapter(sql, conn)
        {
            MissingSchemaAction = MissingSchemaAction.AddWithKey
        };
        var table = new DataTable(name) { Locale = CultureInfo.InvariantCulture };
        if (pageSize == 0)
            adapter.Fill(table);
        else
            adapter.Fill((int)start, pageSize, table);

        // Some Access table layouts do not return a primary key during Fill.
        if (table.PrimaryKey.Length == 0 && primaryKeyNames.Count > 0)
        {
            var columns = primaryKeyNames
                .Where(table.Columns.Contains)
                .Select(column => table.Columns[column]!)
                .ToArray();
            if (columns.Length == primaryKeyNames.Count)
                table.PrimaryKey = columns;
        }
        return new TablePage(table, total, pageIndex, pageSize);
    }

    private static IReadOnlyList<string> FindPrimaryKeys(OleDbConnection conn, string name)
    {
        try
        {
            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Primary_Keys,
                new object?[] { null, null, name });
            if (schema is null) return Array.Empty<string>();
            return schema.Rows.Cast<DataRow>()
                .OrderBy(row => row.Table.Columns.Contains("ORDINAL") &&
                    row["ORDINAL"] != DBNull.Value
                    ? Convert.ToInt32(row["ORDINAL"], CultureInfo.InvariantCulture) : 0)
                .Select(row => Convert.ToString(row["COLUMN_NAME"]) ?? "")
                .Where(column => !string.IsNullOrWhiteSpace(column))
                .ToArray();
        }
        catch (OleDbException)
        {
            // Without a confirmed primary key the table remains read-only.
            return Array.Empty<string>();
        }
    }

    /// <summary>Streams the entire saved table to CSV, not just the displayed page.</summary>
    public long ExportTableCsv(string name, string destination)
    {
        using var conn = Open();
        using var command = new OleDbCommand("SELECT * FROM " + QuoteName(name), conn);
        using var reader = command.ExecuteReader()
            ?? throw new InvalidOperationException("Access did not return a table reader.");
        using var writer = new StreamWriter(destination, false, new UTF8Encoding(true));
        writer.WriteLine(string.Join(",", Enumerable.Range(0, reader.FieldCount)
            .Select(i => CsvCell(reader.GetName(i)))));
        long rows = 0;
        while (reader.Read())
        {
            writer.WriteLine(string.Join(",", Enumerable.Range(0, reader.FieldCount)
                .Select(i => CsvCell(reader.GetValue(i)))));
            rows++;
        }
        return rows;
    }

    private static string CsvCell(object value)
    {
        string raw = value switch
        {
            null or DBNull => "",
            DateTime date => date.ToString("O", CultureInfo.InvariantCulture),
            byte[] bytes => Convert.ToBase64String(bytes),
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
        };
        // Prevent spreadsheet applications from interpreting untrusted cells as formulas.
        if (raw.Length > 0 && "=+-@\t\r".Contains(raw[0])) raw = "'" + raw;
        return "\"" + raw.Replace("\"", "\"\"") + "\"";
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

    /// <summary>
    /// Reads key values and numbered Recordname values from the entire saved
    /// table so copy/paste cannot generate identifiers based only on one page.
    /// This is read-only and does not modify or upload the MDB file.
    /// </summary>
    internal PasteIdentityCatalog ReadPasteIdentityCatalog(DataTable table)
    {
        if (table.PrimaryKey.Length == 0)
            throw new InvalidOperationException("A primary key is required to create new records.");

        var catalog = new PasteIdentityCatalog();
        var keys = table.PrimaryKey;
        var nameColumn = PasteIdentityCatalog.FindRecordName(table);
        var fields = keys.Select(k => k.ColumnName).ToList();
        if (nameColumn is not null &&
            !fields.Contains(nameColumn.ColumnName, StringComparer.OrdinalIgnoreCase))
            fields.Add(nameColumn.ColumnName);

        using var conn = Open();
        using var command = new OleDbCommand("SELECT " +
            string.Join(", ", fields.Select(QuoteName)) + " FROM " +
            QuoteName(table.TableName), conn);
        using var reader = command.ExecuteReader()
            ?? throw new InvalidOperationException("Could not enumerate existing record identifiers.");

        int nameOrdinal = nameColumn is null ? -1 :
            fields.FindIndex(field => string.Equals(field, nameColumn.ColumnName,
                StringComparison.OrdinalIgnoreCase));
        while (reader.Read())
        {
            object?[] values = new object?[keys.Length];
            for (int i = 0; i < keys.Length; i++)
                values[i] = reader.GetValue(i);
            string? recordName = nameOrdinal < 0 || reader.IsDBNull(nameOrdinal) ? null :
                Convert.ToString(reader.GetValue(nameOrdinal), CultureInfo.InvariantCulture);
            catalog.Include(values, recordName);
        }
        catalog.IncludePending(table);
        return catalog;
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
        adapter.Fill(result);
        return result;
    }
}

public sealed record TablePage(DataTable Rows, long TotalRecords, int PageIndex, int PageSize);
