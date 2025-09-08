using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Helper to add common badges to README.
/// </summary>
public sealed class BadgeBuilder {
    private readonly List<(string Text, string Url, string? Title)> _badges;
    internal BadgeBuilder(List<(string Text, string Url, string? Title)> list) { _badges = list; }
    /// <summary>Adds a custom badge label linked to a URL.</summary>
    public BadgeBuilder Custom(string text, string url, string? title = null) { _badges.Add((text, url, title)); return this; }
    /// <summary>Shortcut for a NuGet badge link.</summary>
    public BadgeBuilder NuGet(string? id = null) {
        if (!string.IsNullOrWhiteSpace(id)) _badges.Add(("NuGet", $"https://www.nuget.org/packages/{id}", null));
        else _badges.Add(("NuGet", "https://www.nuget.org/", null));
        return this;
    }
    /// <summary>Adds a CI build badge link. URL required for explicitness.</summary>
    public BadgeBuilder Build(string url) { if (!string.IsNullOrWhiteSpace(url)) _badges.Add(("Build", url, null)); return this; }
    /// <summary>Convenience for GitHub Actions badge link.</summary>
    public BadgeBuilder BuildForGitHub(string owner, string repo, string? workflow = null) {
        var url = string.IsNullOrWhiteSpace(workflow)
            ? $"https://github.com/{owner}/{repo}/actions"
            : $"https://github.com/{owner}/{repo}/actions/workflows/{workflow}";
        _badges.Add(("Build", url, null));
        return this;
    }
    /// <summary>Adds a code coverage badge link. URL required for explicitness.</summary>
    public BadgeBuilder Coverage(string url) { if (!string.IsNullOrWhiteSpace(url)) _badges.Add(("Coverage", url, null)); return this; }
    /// <summary>Convenience for Codecov coverage link.</summary>
    public BadgeBuilder CoverageCodecov(string owner, string repo, string? branch = null) {
        var url = string.IsNullOrWhiteSpace(branch)
            ? $"https://codecov.io/gh/{owner}/{repo}"
            : $"https://codecov.io/gh/{owner}/{repo}/branch/{branch}";
        _badges.Add(("Coverage", url, null));
        return this;
    }
}
