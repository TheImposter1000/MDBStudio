using Microsoft.Win32;

namespace MDBStudio;

/// <summary>
/// Registers MDB Studio as an available Windows Open With handler, without
/// stealing the current default app association for Access files.
/// Windows 10/11 require users to choose their own default in system UI.
/// </summary>
public static class WindowsFileAssociation
{
    private const string ProgId = "MDBStudio.AccessDatabase.1";
    private const string RegisteredName = "MDB Studio";
    private const string CapabilitiesPath = @"Software\MDBStudio\Capabilities";

    public static string? DatabasePathFromArguments(string[] args)
    {
        if (args.Length != 1 || string.IsNullOrWhiteSpace(args[0])) return null;
        string extension = Path.GetExtension(args[0]);
        if (!extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase) &&
            !extension.Equals(".accdb", StringComparison.OrdinalIgnoreCase))
            return null;
        return Path.GetFullPath(args[0]);
    }

    public static string OpenCommand(string executablePath)
    {
        if (string.IsNullOrWhiteSpace(executablePath))
            throw new ArgumentException("The application executable is missing.", nameof(executablePath));
        return "\"" + Path.GetFullPath(executablePath) + "\" \"%1\"";
    }

    public static void RegisterForCurrentUser(string executablePath)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("File associations are supported on Windows.");

        string fullPath = Path.GetFullPath(executablePath);
        if (!string.Equals(Path.GetFileName(fullPath), "MDBStudio.exe", StringComparison.OrdinalIgnoreCase) ||
            !File.Exists(fullPath))
            throw new InvalidOperationException("Only an installed MDBStudio.exe can register file associations.");

        const string classes = @"Software\Classes\";
        using (var prog = Registry.CurrentUser.CreateSubKey(classes + ProgId))
        {
            if (prog is null) throw new InvalidOperationException("Cannot register MDB Studio for this Windows user.");
            prog.SetValue("", "MDB Studio Access Database");
        }
        using (var icon = Registry.CurrentUser.CreateSubKey(classes + ProgId + @"\DefaultIcon"))
            icon?.SetValue("", "\"" + fullPath + "\",0");
        using (var command = Registry.CurrentUser.CreateSubKey(classes + ProgId + @"\shell\open\command"))
            command?.SetValue("", OpenCommand(fullPath));

        foreach (string extension in new[] { ".mdb", ".accdb" })
        {
            using var openWith = Registry.CurrentUser.CreateSubKey(
                classes + extension + @"\OpenWithProgids");
            openWith?.SetValue(ProgId, "", RegistryValueKind.String);
        }

        using (var app = Registry.CurrentUser.CreateSubKey(
            @"Software\Microsoft\Windows\CurrentVersion\App Paths\MDBStudio.exe"))
            app?.SetValue("", fullPath);
        using (var capabilities = Registry.CurrentUser.CreateSubKey(CapabilitiesPath))
        {
            capabilities?.SetValue("ApplicationName", RegisteredName);
            capabilities?.SetValue("ApplicationDescription",
                "Browse, edit, import and create Microsoft Access MDB and ACCDB databases.");
        }
        using (var formats = Registry.CurrentUser.CreateSubKey(
            CapabilitiesPath + @"\FileAssociations"))
        {
            formats?.SetValue(".mdb", ProgId);
            formats?.SetValue(".accdb", ProgId);
        }
        using (var installed = Registry.CurrentUser.CreateSubKey(@"Software\RegisteredApplications"))
            installed?.SetValue(RegisteredName, CapabilitiesPath);
    }

    public const string DefaultAppsSettingsUri =
        "ms-settings:defaultapps?registeredAppUser=MDB%20Studio";
}
