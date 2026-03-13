namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Options controlling Chart.js rendering when used in a WebView/browser context.
/// Generic charts are authored as fenced code blocks named <c>chart</c> containing JSON.
/// IntelligenceX/chat hosts can opt into the <c>ix-chart</c> alias through <see cref="MarkdownRendererIntelligenceXAdapter"/>
/// or the chat preset family.
/// </summary>
public sealed class ChartOptions {
    /// <summary>Enable Chart.js support. Default: false.</summary>
    public bool Enabled { get; set; } = false;

    /// <summary>
    /// Script URL for Chart.js (UMD). Default points at jsDelivr.
    /// Hosts can override this to use a local asset for offline scenarios.
    /// </summary>
    public string ScriptUrl { get; set; } = "https://cdn.jsdelivr.net/npm/chart.js/dist/chart.umd.min.js";
}

