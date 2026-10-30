using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace MDBStudio;

/// <summary>
/// Uses the public official GitHub Releases API. No arbitrary update source or
/// background installation. Installs only after user confirmation and SHA256 check.
/// </summary>
public static class UpdateChecker
{
    public const string CurrentReleaseTag = "v1.0.2";
    private const string Repository = "TheImposter1000/MDBStudio";
    private static readonly Uri ApiUrl =
        new("https://api.github.com/repos/" + Repository + "/releases?per_page=100");

    private static HttpClient Client()
    {
        var client = new HttpClient { Timeout = TimeSpan.FromMinutes(8) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("MDBStudio/1.0.2");
        client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        return client;
    }

    private static Uri? AssetUrl(JsonElement release, string fileName, string tag)
    {
        if (!release.TryGetProperty("assets", out var assets)) return null;
        foreach (var asset in assets.EnumerateArray())
        {
            if (!asset.TryGetProperty("name", out var name) || name.GetString() != fileName ||
                !asset.TryGetProperty("browser_download_url", out var urlElement) ||
                !Uri.TryCreate(urlElement.GetString(), UriKind.Absolute, out var url)) continue;
            if (url.Scheme == Uri.UriSchemeHttps &&
                string.Equals(url.Host, "github.com", StringComparison.OrdinalIgnoreCase) &&
                url.AbsolutePath.StartsWith(
                    "/" + Repository + "/releases/download/" + tag + "/",
                    StringComparison.OrdinalIgnoreCase))
                return url;
        }
        return null;
    }

    public static async Task<UpdateResult> CheckAsync()
    {
        using var client = Client();
        using var response = await client.GetAsync(ApiUrl);
        response.EnsureSuccessStatusCode();
        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
        if (!ReleaseVersion.TryParse(CurrentReleaseTag, out var current))
            throw new InvalidOperationException("Invalid installed application version.");

        ReleaseVersion? latest = null;
        UpdateResult? found = null;
        foreach (var release in document.RootElement.EnumerateArray())
        {
            if (release.TryGetProperty("draft", out var draft) && draft.GetBoolean()) continue;
            if (release.TryGetProperty("prerelease", out var prerelease) && prerelease.GetBoolean()) continue;
            if (!release.TryGetProperty("tag_name", out var tagField) ||
                !ReleaseVersion.TryParse(tagField.GetString(), out var candidate))
                continue;
            string tag = tagField.GetString()!;
            if (!release.TryGetProperty("html_url", out var pageField) ||
                !Uri.TryCreate(pageField.GetString(), UriKind.Absolute, out var page) ||
                page.Scheme != Uri.UriSchemeHttps ||
                !string.Equals(page.Host, "github.com", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(page.AbsolutePath,
                    "/" + Repository + "/releases/tag/" + tag, StringComparison.OrdinalIgnoreCase))
                continue;

            if (latest is null || candidate.CompareTo(latest) > 0)
            {
                latest = candidate;
                string zipName = $"MDBStudio-Windows-x64-{tag}.zip";
                found = new UpdateResult(CurrentReleaseTag, tag, page,
                    AssetUrl(release, zipName, tag), AssetUrl(release, zipName + ".sha256", tag),
                    candidate.CompareTo(current) > 0);
            }
        }
        return found ?? throw new InvalidOperationException("No supported MDB Studio releases found.");
    }

    private static async Task DownloadAsync(HttpClient client, Uri url, string destination, long maxBytes)
    {
        using var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();
        if (response.Content.Headers.ContentLength is long size && size > maxBytes)
            throw new InvalidOperationException("The update package exceeds the size limit.");

        await using var source = await response.Content.ReadAsStreamAsync();
        await using var output = File.Create(destination);
        byte[] buffer = new byte[128 * 1024];
        long total = 0;
        int count;
        while ((count = await source.ReadAsync(buffer)) > 0)
        {
            total += count;
            if (total > maxBytes)
                throw new InvalidOperationException("The downloaded update is unexpectedly large.");
            await output.WriteAsync(buffer.AsMemory(0, count));
        }
    }

    public static async Task<StagedUpdate> DownloadAndVerifyAsync(UpdateResult release)
    {
        if (!release.UpdateAvailable || release.ZipUrl is null || release.ChecksumUrl is null)
            throw new InvalidOperationException(
                "This version does not have a verified in-app update package. Open its GitHub release page instead.");

        string work = Path.Combine(Path.GetTempPath(), "MDBStudio-Update-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(work);
        string zip = Path.Combine(work, "update.zip");
        string checksumFile = Path.Combine(work, "update.sha256");
        try
        {
            using var client = Client();
            await DownloadAsync(client, release.ChecksumUrl, checksumFile, 1024);
            await DownloadAsync(client, release.ZipUrl, zip, 700_000_000);
            string checksumText = await File.ReadAllTextAsync(checksumFile);
            string expectedZipName = $"MDBStudio-Windows-x64-{release.LatestTag}.zip";
            var match = Regex.Match(checksumText,
                @"^\s*([0-9a-fA-F]{64})\s+\*?(.+?)\s*$", RegexOptions.CultureInvariant);
            if (!match.Success || !string.Equals(match.Groups[2].Value, expectedZipName, StringComparison.Ordinal))
                throw new InvalidOperationException("The release checksum manifest did not match the package.");

            string expected = match.Groups[1].Value.ToUpperInvariant();
            await using var stream = File.OpenRead(zip);
            string actual = Convert.ToHexString(await SHA256.HashDataAsync(stream));
            if (!string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Downloaded update failed checksum verification.");

            string updater = Path.Combine(AppContext.BaseDirectory, "MDBStudio.Updater.exe");
            if (!File.Exists(updater))
                throw new InvalidOperationException(
                    "The installed MDB Studio has no updater helper. Install v1.0.2 once from GitHub to enable future in-app updates.");
            string tempUpdater = Path.Combine(work, "MDBStudio.Updater.exe");
            File.Copy(updater, tempUpdater);
            return new StagedUpdate(work, zip, expected, tempUpdater);
        }
        catch
        {
            try { Directory.Delete(work, recursive: true); } catch { }
            throw;
        }
    }

    internal sealed record ReleaseVersion(Version Number, int Preview) : IComparable<ReleaseVersion>
    {
        private static readonly Regex Pattern = new(
            @"^v(?<number>\d+\.\d+\.\d+)(?:-preview\.(?<preview>\d+))?$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public static bool TryParse(string? tag, out ReleaseVersion version)
        {
            version = null!;
            var m = Pattern.Match(tag ?? "");
            if (!m.Success || !Version.TryParse(m.Groups["number"].Value, out var number)) return false;
            int preview = int.MaxValue;
            if (m.Groups["preview"].Success && !int.TryParse(m.Groups["preview"].Value, out preview))
                return false;
            version = new ReleaseVersion(number, preview);
            return true;
        }

        public int CompareTo(ReleaseVersion? other)
        {
            if (other is null) return 1;
            int result = Number.CompareTo(other.Number);
            return result != 0 ? result : Preview.CompareTo(other.Preview);
        }
    }
}
public sealed record UpdateResult(string CurrentTag, string LatestTag, Uri LatestUrl,
    Uri? ZipUrl, Uri? ChecksumUrl, bool UpdateAvailable);
public sealed record StagedUpdate(string WorkDirectory, string ArchivePath, string Sha256, string UpdaterPath);
