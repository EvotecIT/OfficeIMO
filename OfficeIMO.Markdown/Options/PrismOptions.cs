namespace OfficeIMO.Markdown;

/// <summary>
/// Options for Prism.js syntax highlighting.
/// </summary>
public sealed class MarkdownPrismOptions {
    /// <summary>Enable Prism highlighting. Default: false.</summary>
    public bool Enabled { get; set; } = false;
    /// <summary>Theme stylesheet. Default: <see cref="MarkdownPrismTheme.Prism"/>. Use <see cref="MarkdownPrismTheme.GithubAuto"/> to pair with GithubAuto markdown style.</summary>
    public MarkdownPrismTheme Theme { get; set; } = MarkdownPrismTheme.Prism;
    /// <summary>Language ids to include. If empty, a small default set is used.</summary>
    public List<string> Languages { get; } = new() { "markup", "csharp", "bash", "powershell", "json", "yaml", "markdown" };
    /// <summary>Plugin ids to include (e.g., "line-numbers", "copy-to-clipboard").</summary>
    public List<string> Plugins { get; } = new();
    /// <summary>CDN base URL. Default: jsDelivr.</summary>
    public string CdnBase { get; set; } = "https://cdn.jsdelivr.net/npm/prismjs@1.29.0";

    internal MarkdownPrismOptions CloneForRender() {
        var clone = new MarkdownPrismOptions {
            Enabled = Enabled,
            Theme = Theme,
            CdnBase = CdnBase
        };
        clone.Languages.Clear();
        clone.Languages.AddRange(Languages);
        clone.Plugins.AddRange(Plugins);
        return clone;
    }
}
