using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace MDBStudio;

/// <summary>
/// Creates a completely new Microsoft Access file without launching Access.exe.
/// Uses the installed 64-bit DAO or ADOX Access database components.
/// The finished database is verified and moved into place without overwriting
/// existing MDB/ACCDB files, even when another process creates the same name.
/// </summary>
public static class AccessDatabaseCreator
{
    private const string GeneralLocale = ";LANGID=0x0409;CP=1252;COUNTRY=0";
    private const int DaoJet4 = 64;
    private const int DaoAccess12 = 128;

    public static string ValidateNewPath(string filename)
    {
        if (string.IsNullOrWhiteSpace(filename))
            throw new ArgumentException("Choose a name for the new database.", nameof(filename));
        string path = Path.GetFullPath(filename);
        string extension = Path.GetExtension(path);
        if (!extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".accdb", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Create an Access .mdb or .accdb database.");
        if (!Directory.Exists(Path.GetDirectoryName(path)))
            throw new DirectoryNotFoundException("Choose an existing folder for the new database.");
        if (File.Exists(path) || Directory.Exists(path))
            throw new IOException("A file or folder already exists at this location. Choose a different name. MDB Studio never overwrites an existing database when creating one.");
        return path;
    }

    public static string Create(string filename)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Creating Access databases requires Windows and the Microsoft Access Database Engine.");

        string destination = ValidateNewPath(filename);
        string extension = Path.GetExtension(destination);
        string tempFile = Path.Combine(Path.GetDirectoryName(destination)!,
            "MDBStudio-New-" + Guid.NewGuid().ToString("N") + extension);
        var errors = new List<string>();
        try
        {
            // DAO is the native Access creation API. The 64/128 constants mean
            // Jet 4 (.mdb) and Access 12 (.accdb), respectively.
            try
            {
                CreateWithDao(tempFile, extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase)
                    ? DaoJet4 : DaoAccess12);
            }
            catch (Exception ex) when (ex is COMException or InvalidOperationException or
                MissingMethodException or TypeLoadException or System.Reflection.TargetInvocationException)
            {
                errors.Add("DAO: " + ex.Message);
                if (File.Exists(tempFile)) File.Delete(tempFile);
                try { CreateWithAdox(tempFile, extension); }
                catch (Exception fallback) when (fallback is COMException or InvalidOperationException or
                    MissingMethodException or TypeLoadException or System.Reflection.TargetInvocationException)
                {
                    errors.Add("ADOX: " + fallback.Message);
                    throw new InvalidOperationException(
                        "Cannot create a new Access database with the installed 64-bit database components. " +
                        "Install or repair the matching 64-bit Microsoft Access Database Engine (DAO/ADOX). " +
                        "The full Microsoft Access application is not required.\n\n" +
                        string.Join("\n", errors), fallback);
                }
            }

            if (!File.Exists(tempFile))
                throw new InvalidOperationException("The Access database engine returned without creating a database file.");
            // Verify it is readable before placing the file at the requested name.
            var verification = new AccessSession(tempFile);
            _ = verification.ListTables();
            // File.Move without overwrite is atomic on the same volume.
            File.Move(tempFile, destination);
            return destination;
        }
        finally
        {
            // Do not leave a partially generated or locked temporary file behind.
            if (File.Exists(tempFile))
            {
                try { File.Delete(tempFile); } catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    private static void CreateWithDao(string file, int version)
    {
        Type engineType = Type.GetTypeFromProgID("DAO.DBEngine.120")
            ?? throw new InvalidOperationException("The DAO Access Database Engine is not registered.");
        object engine = Activator.CreateInstance(engineType)
            ?? throw new InvalidOperationException("Unable to instantiate the DAO Access Database Engine.");
        object? database = null;
        try
        {
            dynamic dao = engine;
            database = dao.CreateDatabase(file, GeneralLocale, version);
        }
        finally
        {
            if (database is not null)
            {
                try { ((dynamic)database).Close(); } finally { ReleaseCom(database); }
            }
            ReleaseCom(engine);
        }
    }

    private static void CreateWithAdox(string file, string extension)
    {
        Type catalogType = Type.GetTypeFromProgID("ADOX.Catalog")
            ?? throw new InvalidOperationException("The ADOX Catalog component is not registered.");
        object catalog = Activator.CreateInstance(catalogType)
            ?? throw new InvalidOperationException("Unable to instantiate ADOX Catalog.");
        object? connection = null;
        try
        {
            var builder = new OleDbConnectionStringBuilder
            {
                Provider = "Microsoft.ACE.OLEDB.16.0",
                DataSource = file
            };
            builder["Jet OLEDB:Engine Type"] =
                extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase) ? 5 : 6;
            dynamic adox = catalog;
            adox.Create(builder.ConnectionString);
            connection = adox.ActiveConnection;
        }
        finally
        {
            if (connection is not null)
            {
                try { ((dynamic)connection).Close(); } finally { ReleaseCom(connection); }
            }
            ReleaseCom(catalog);
        }
    }

    private static void ReleaseCom(object instance)
    {
        if (Marshal.IsComObject(instance))
            Marshal.FinalReleaseComObject(instance);
    }
}
