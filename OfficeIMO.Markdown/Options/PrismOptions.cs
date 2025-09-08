using System.Collections.Generic;

namespace OfficeIMO.Markdown;

/// <summary>
/// Options for Prism.js syntax highlighting.
/// </summary>
public sealed class PrismOptions {
    /// <summary>Enable Prism highlighting. Default: false.</summary>
    public bool Enabled { get; set; } = false;
    /// <summary>Theme stylesheet. Default: <see cref="PrismTheme.Prism"/>. Use <see cref="PrismTheme.GithubAuto"/> to pair with GithubAuto markdown style.</summary>
    public PrismTheme Theme { get; set; } = PrismTheme.Prism;
    /// <summary>Language ids to include. If empty, a small default set is used.</summary>
    public List<string> Languages { get; } = new() { "markup", "csharp", "bash", "powershell", "json", "yaml", "markdown" };
    /// <summary>Plugin ids to include (e.g., "line-numbers", "copy-to-clipboard").</summary>
    public List<string> Plugins { get; } = new();
    /// <summary>CDN base URL. Default: jsDelivr.</summary>
    public string CdnBase { get; set; } = "https://cdn.jsdelivr.net/npm/prismjs@1.29.0";
}
