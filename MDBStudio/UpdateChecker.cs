using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace MDBStudio;

/// <summary>
/// Explicit, read-only update check against GitHub's public release metadata.
/// Never downloads, executes, or installs releases automatically.
/// </summary>
public static class UpdateChecker
{
    public const string CurrentReleaseTag = "v0.1.2-preview.1";
    private const string Repository = "TheImposter1000/MDBStudio";
    private static readonly Uri ApiUrl =
        new("https://api.github.com/repos/" + Repository + "/releases?per_page=100");

    public static async Task<UpdateResult> CheckAsync()
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("MDBStudio/0.1.2");
        client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
        using var response = await client.GetAsync(ApiUrl);
        response.EnsureSuccessStatusCode();
        using var document = JsonDocument.Parse(await response.Content.ReadAsStringAsync());

        if (!ReleaseVersion.TryParse(CurrentReleaseTag, out var current))
            throw new InvalidOperationException("Application release version is invalid.");

        ReleaseVersion? latestVersion = null;
        string? latestTag = null;
        Uri? latestUrl = null;
        foreach (var item in document.RootElement.EnumerateArray())
        {
            if (item.TryGetProperty("draft", out var draft) && draft.GetBoolean())
                continue;
            if (!item.TryGetProperty("tag_name", out var tagField) ||
                !ReleaseVersion.TryParse(tagField.GetString(), out var candidate))
                continue;
            if (!item.TryGetProperty("html_url", out var linkField) ||
                !Uri.TryCreate(linkField.GetString(), UriKind.Absolute, out var link))
                continue;
            if (link.Scheme != Uri.UriSchemeHttps ||
                !string.Equals(link.Host, "github.com", StringComparison.OrdinalIgnoreCase) ||
                !link.AbsolutePath.StartsWith("/" + Repository + "/releases/tag/", StringComparison.OrdinalIgnoreCase))
                continue;

            if (latestVersion is null || candidate.CompareTo(latestVersion) > 0)
            {
                latestVersion = candidate;
                latestTag = tagField.GetString();
                latestUrl = link;
            }
        }

        if (latestVersion is null || latestTag is null || latestUrl is null)
            throw new InvalidOperationException("No compatible MDB Studio releases are published on GitHub.");

        return new UpdateResult(CurrentReleaseTag, latestTag, latestUrl,
            latestVersion.CompareTo(current) > 0);
    }

    internal sealed record ReleaseVersion(Version Number, int Preview) : IComparable<ReleaseVersion>
    {
        private static readonly Regex TagPattern = new(
            @"^v(?<number>\d+\.\d+\.\d+)(?:-preview\.(?<preview>\d+))?$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public static bool TryParse(string? tag, out ReleaseVersion version)
        {
            version = null!;
            var match = TagPattern.Match(tag ?? "");
            if (!match.Success ||
                !Version.TryParse(match.Groups["number"].Value, out var number))
                return false;

            // Stable releases take precedence over previews of the same version.
            int preview = int.MaxValue;
            if (match.Groups["preview"].Success &&
                !int.TryParse(match.Groups["preview"].Value, out preview))
                return false;
            version = new ReleaseVersion(number, preview);
            return true;
        }

        public int CompareTo(ReleaseVersion? other)
        {
            if (other is null) return 1;
            var compared = Number.CompareTo(other.Number);
            return compared != 0 ? compared : Preview.CompareTo(other.Preview);
        }
    }
}

public sealed record UpdateResult(string CurrentTag, string LatestTag, Uri LatestUrl, bool UpdateAvailable);
