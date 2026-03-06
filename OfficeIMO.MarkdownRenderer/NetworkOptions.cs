namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Options controlling vis-network rendering when used in a WebView/browser context.
/// Networks are authored as fenced code blocks named <c>ix-network</c>, <c>network</c>, or <c>visnetwork</c>
/// containing JSON with <c>nodes</c>, <c>edges</c>, and optional <c>options</c>.
/// </summary>
public sealed class NetworkOptions {
    /// <summary>Enable vis-network support. Default: false.</summary>
    public bool Enabled { get; set; } = false;

    /// <summary>
    /// Script URL for the standalone vis-network UMD bundle. Default points at the official unpkg standalone build.
    /// Hosts can override this to use a local asset for offline scenarios.
    /// </summary>
    public string ScriptUrl { get; set; } = "https://unpkg.com/vis-network/standalone/umd/vis-network.min.js";

    /// <summary>
    /// Stylesheet URL for vis-network. Hosts can override this to use a local asset for offline scenarios.
    /// </summary>
    public string CssUrl { get; set; } = "https://unpkg.com/vis-network/styles/vis-network.min.css";
}
