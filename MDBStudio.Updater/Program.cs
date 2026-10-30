using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace MDBStudio.Updater;

/// <summary>
/// Small, independently running updater. The main app copies this self-contained
/// executable to the temporary update directory before closing. All paths and
/// downloaded bytes are checked before any installed files are changed.
/// </summary>
internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if (args.Length == 1 && args[0] == "--self-test")
            {
                SelfTest();
                return 0;
            }
            if (args.Length != 5 || args[0] != "--apply" ||
                !int.TryParse(args[1], out int processId) || processId <= 0)
                throw new InvalidOperationException("Invalid updater launch. Run updates from MDB Studio.");

            string target = Path.GetFullPath(args[2]);
            string archive = Path.GetFullPath(args[3]);
            string expected = args[4];
            if (target == Path.GetPathRoot(target) || !File.Exists(Path.Combine(target, "MDBStudio.exe")))
                throw new InvalidOperationException("Target is not a valid MDB Studio application folder.");

            WaitForExit(processId);
            Apply(target, archive, expected);
            _ = Process.Start(new ProcessStartInfo(Path.Combine(target, "MDBStudio.exe"))
            {
                UseShellExecute = true,
                WorkingDirectory = target
            }) ?? throw new InvalidOperationException("Update installed but MDB Studio did not relaunch.");
            return 0;
        }
        catch (Exception ex)
        {
            MessageBox.Show("The update could not be completed.\n\n" + ex.Message +
                "\n\nYour previous application files have been restored where possible. " +
                "Do not disable Windows Defender to complete an update.",
                "MDB Studio update failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }

    private static void WaitForExit(int processId)
    {
        try
        {
            using Process p = Process.GetProcessById(processId);
            if (!p.WaitForExit(90_000))
                throw new InvalidOperationException("Close MDB Studio and retry the update.");
        }
        catch (ArgumentException)
        {
            // The application already exited.
        }
    }

    private static void Verify(string archive, string sha)
    {
        if (sha.Length != 64 || !sha.All(Uri.IsHexDigit))
            throw new InvalidOperationException("The download's SHA256 checksum is invalid.");
        using var stream = File.OpenRead(archive);
        var actual = Convert.ToHexString(SHA256.HashData(stream));
        if (!string.Equals(actual, sha, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The downloaded release failed checksum verification. No files were changed.");
    }

    internal static void Apply(string target, string archive, string expectedHash, bool failAfterFirstWrite = false)
    {
        Verify(archive, expectedHash);
        string work = Path.Combine(Path.GetTempPath(), "MDBStudio-Apply-" + Guid.NewGuid().ToString("N"));
        string stage = Path.Combine(work, "stage");
        string backups = Path.Combine(work, "backup");
        Directory.CreateDirectory(stage);
        Directory.CreateDirectory(backups);
        var overwritten = new List<string>();
        var newlyCreated = new List<string>();
        try
        {
            if ((File.GetAttributes(target) & FileAttributes.ReparsePoint) != 0)
                throw new InvalidOperationException("Updating through a redirected or linked folder is not supported.");

            using (var zip = ZipFile.OpenRead(archive))
            {
                if (zip.Entries.Count > 3000)
                    throw new InvalidOperationException("The release contains too many files.");
                long expandedSize = 0;
                foreach (var entry in zip.Entries)
                {
                    expandedSize = checked(expandedSize + entry.Length);
                    if (expandedSize > 1_000_000_000L)
                        throw new InvalidOperationException("The release exceeds the safe size limit.");
                    string relative = entry.FullName.Replace('/', Path.DirectorySeparatorChar);
                    if (relative.StartsWith(Path.DirectorySeparatorChar) || Path.IsPathRooted(relative))
                        throw new InvalidOperationException("The release contains an unsafe absolute path.");
                    string output = Path.GetFullPath(Path.Combine(stage, relative));
                    if (!output.StartsWith(stage + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The release contains an unsafe path.");
                    if (((entry.ExternalAttributes >> 16) & 0xF000) == 0xA000)
                        throw new InvalidOperationException("Symbolic links are not accepted inside releases.");
                    if (entry.Name.Length == 0) continue;
                    var extension = Path.GetExtension(entry.Name);
                    if (new[] { ".mdb", ".accdb", ".ldb", ".laccdb", ".bat", ".cmd", ".ps1", ".vbs", ".js" }
                        .Contains(extension, StringComparer.OrdinalIgnoreCase))
                        throw new InvalidOperationException("The release contains an unsupported data file or script.");
                    Directory.CreateDirectory(Path.GetDirectoryName(output)!);
                    entry.ExtractToFile(output, overwrite: false);
                }
            }

            if (!File.Exists(Path.Combine(stage, "MDBStudio.exe")) ||
                !File.Exists(Path.Combine(stage, "MDBStudio.Updater.exe")))
                throw new InvalidOperationException("The release does not contain MDB Studio and its updater.");

            var stagedFiles = Directory.GetFiles(stage, "*", SearchOption.AllDirectories);
            foreach (var staged in stagedFiles)
            {
                var relative = Path.GetRelativePath(stage, staged);
                var dest = Path.GetFullPath(Path.Combine(target, relative));
                if (!dest.StartsWith(target + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Invalid release destination path.");
                Directory.CreateDirectory(Path.GetDirectoryName(dest)!);

                if (File.Exists(dest))
                {
                    string backup = Path.Combine(backups, relative);
                    Directory.CreateDirectory(Path.GetDirectoryName(backup)!);
                    File.Copy(dest, backup, true);
                    overwritten.Add(relative);
                }
                else newlyCreated.Add(relative);

                File.Copy(staged, dest, true);
                if (failAfterFirstWrite)
                    throw new InvalidOperationException("Simulated update failure.");
            }
        }
        catch
        {
            // Restore the pre-update application if any replacement failed.
            foreach (var relative in overwritten.AsEnumerable().Reverse())
            {
                string backup = Path.Combine(backups, relative);
                if (File.Exists(backup)) File.Copy(backup, Path.Combine(target, relative), true);
            }
            foreach (var relative in newlyCreated.AsEnumerable().Reverse())
            {
                string path = Path.Combine(target, relative);
                if (File.Exists(path)) File.Delete(path);
            }
            throw;
        }
        finally
        {
            try { Directory.Delete(work, recursive: true); } catch { /* best effort */ }
        }
    }

    private static void SelfTest()
    {
        string root = Path.Combine(Path.GetTempPath(), "MDBStudio-SelfTest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string target = Path.Combine(root, "installed");
            Directory.CreateDirectory(target);
            File.WriteAllText(Path.Combine(target, "MDBStudio.exe"), "OLD");
            string archive = Path.Combine(root, "release.zip");
            using (var zip = ZipFile.Open(archive, ZipArchiveMode.Create))
            {
                var a = zip.CreateEntry("MDBStudio.exe");
                using (var w = new StreamWriter(a.Open())) w.Write("NEW");
                var b = zip.CreateEntry("MDBStudio.Updater.exe");
                using (var w = new StreamWriter(b.Open())) w.Write("UPDATER");
            }
            string sha = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(archive)));
            try
            {
                Apply(target, archive, sha, failAfterFirstWrite: true);
                throw new Exception("Rollback test did not fail.");
            }
            catch (InvalidOperationException ex) when (ex.Message == "Simulated update failure.") { }
            if (File.ReadAllText(Path.Combine(target, "MDBStudio.exe")) != "OLD")
                throw new Exception("Update rollback failed.");
            Apply(target, archive, sha);
            if (File.ReadAllText(Path.Combine(target, "MDBStudio.exe")) != "NEW" ||
                !File.Exists(Path.Combine(target, "MDBStudio.Updater.exe")))
                throw new Exception("Updater installation test failed.");
            try
            {
                Apply(target, archive, new string('0', 64));
                throw new Exception("Invalid hash should be rejected.");
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("checksum")) { }
            Console.WriteLine("Updater verification, replacement and rollback tests passed.");
        }
        finally
        {
            try { Directory.Delete(root, true); } catch { }
        }
    }
}
