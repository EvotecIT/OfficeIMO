namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Options controlling Mermaid rendering when used in a WebView/browser context.
/// Mermaid transforms <c>```mermaid</c> fenced code blocks into diagrams at runtime.
/// </summary>
public sealed class MermaidOptions {
    /// <summary>Enable Mermaid support. Default: true.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
     /// ESM module URL for Mermaid. Default points at Mermaid v11 on jsDelivr.
     /// </summary>
    public string EsmModuleUrl { get; set; } = "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.esm.min.mjs";

    /// <summary>
    /// Script URL for Mermaid (non-module). This is useful for offline bundling/hosting scenarios.
    /// Default points at Mermaid v11 on jsDelivr.
    /// </summary>
    public string ScriptUrl { get; set; } = "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js";

    /// <summary>Mermaid theme name to use in light mode. Default: "default".</summary>
    public string LightTheme { get; set; } = "default";

    /// <summary>Mermaid theme name to use in dark mode. Default: "dark".</summary>
    public string DarkTheme { get; set; } = "dark";

    /// <summary>
    /// When true, <see cref="MarkdownRenderer.RenderBodyHtml"/> adds hashes to Mermaid blocks so the incremental updater
    /// can preserve already-rendered SVGs when content is updated.
    /// Default: true.
    /// </summary>
    public bool EnableHashCaching { get; set; } = true;
}

