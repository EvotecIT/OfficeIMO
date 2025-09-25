namespace OfficeIMO.Markdown;

/// <summary>
/// Helper to add common badges to README.
/// </summary>
public sealed class BadgeBuilder {
    private readonly List<(string Text, string Url, string? Title)> _badges;
    private readonly List<(string Alt, string LinkUrl, string ImageUrl)> _imageBadges;
    internal BadgeBuilder(List<(string Text, string Url, string? Title)> list, List<(string Alt, string LinkUrl, string ImageUrl)> images) { _badges = list; _imageBadges = images; }
    /// <summary>Adds a custom badge label linked to a URL.</summary>
    public BadgeBuilder Custom(string text, string url, string? title = null) { _badges.Add((text, url, title)); return this; }
    /// <summary>Adds a NuGet badge using Shields.io.</summary>
    public BadgeBuilder NuGet(string id) {
        if (string.IsNullOrWhiteSpace(id)) return this;
        var link = $"https://www.nuget.org/packages/{id}";
        var img = $"https://img.shields.io/nuget/v/{id}?label=NuGet";
        _imageBadges.Add(("NuGet", link, img));
        return this;
    }
    /// <summary>Adds a CI build badge link (explicit URL only, no dynamic image).</summary>
    public BadgeBuilder Build(string url) { if (!string.IsNullOrWhiteSpace(url)) _badges.Add(("Build", url, null)); return this; }
    /// <summary>Convenience for GitHub Actions badge with Shields.io (workflow file required for dynamic status image).</summary>
    public BadgeBuilder BuildForGitHub(string owner, string repo, string? workflow = null, string? branch = null) {
        var link = string.IsNullOrWhiteSpace(workflow)
            ? $"https://github.com/{owner}/{repo}/actions"
            : $"https://github.com/{owner}/{repo}/actions/workflows/{workflow}";
        if (!string.IsNullOrWhiteSpace(workflow)) {
            var img = $"https://img.shields.io/github/actions/workflow/status/{owner}/{repo}/{workflow}?label=Build" + (string.IsNullOrWhiteSpace(branch) ? "" : $"&branch={branch}");
            _imageBadges.Add(("Build", link, img));
        } else {
            _badges.Add(("Build", link, null));
        }
        return this;
    }
    /// <summary>Adds a code coverage badge using Shields.io for Codecov.</summary>
    public BadgeBuilder Coverage(string url) { if (!string.IsNullOrWhiteSpace(url)) _badges.Add(("Coverage", url, null)); return this; }
    /// <summary>Adds a Codecov coverage badge and link for the given repo/branch.</summary>
    /// <param name="owner">GitHub owner or org.</param>
    /// <param name="repo">Repository name.</param>
    /// <param name="branch">Optional branch for the badge.</param>
    public BadgeBuilder CoverageCodecov(string owner, string repo, string? branch = null) {
        var link = string.IsNullOrWhiteSpace(branch)
            ? $"https://codecov.io/gh/{owner}/{repo}"
            : $"https://codecov.io/gh/{owner}/{repo}/branch/{branch}";
        var img = $"https://img.shields.io/codecov/c/github/{owner}/{repo}?label=coverage" + (string.IsNullOrWhiteSpace(branch) ? "" : $"&branch={branch}");
        _imageBadges.Add(("Coverage", link, img));
        return this;
    }
}
