namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Options controlling math rendering (KaTeX) when used in a WebView/browser context.
/// Math is authored as plain text delimiters (for example <c>$...$</c> or <c>$$...$$</c>) and rendered client-side.
/// </summary>
public sealed class MathOptions {
    /// <summary>Enable math rendering support. Default: true.</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>
    /// CSS URL for KaTeX. Default points at jsDelivr.
    /// Hosts can override this to use local assets for offline scenarios.
    /// </summary>
    public string CssUrl { get; set; } = "https://cdn.jsdelivr.net/npm/katex/dist/katex.min.css";

    /// <summary>
    /// Script URL for KaTeX. Default points at jsDelivr.
    /// </summary>
    public string ScriptUrl { get; set; } = "https://cdn.jsdelivr.net/npm/katex/dist/katex.min.js";

    /// <summary>
    /// Script URL for KaTeX auto-render extension. Default points at jsDelivr.
    /// </summary>
    public string AutoRenderScriptUrl { get; set; } = "https://cdn.jsdelivr.net/npm/katex/dist/contrib/auto-render.min.js";

    /// <summary>Enable <c>$...$</c> inline math. Default: true.</summary>
    public bool EnableDollarInline { get; set; } = true;

    /// <summary>Enable <c>$$...$$</c> display math. Default: true.</summary>
    public bool EnableDollarDisplay { get; set; } = true;

    /// <summary>Enable <c>\\(...\\)</c> inline math. Default: true.</summary>
    public bool EnableParenInline { get; set; } = true;

    /// <summary>Enable <c>\\[...\\]</c> display math. Default: true.</summary>
    public bool EnableBracketDisplay { get; set; } = true;

    /// <summary>
    /// When true, fenced code blocks named <c>math</c>/<c>latex</c> are converted into display-math nodes before rendering,
    /// so authors can write:
    /// <code>```math\nx^2\n```</code>
    /// Default: true.
    /// </summary>
    public bool EnableFencedMathBlocks { get; set; } = true;

    /// <summary>
    /// Fence languages treated as display math when <see cref="EnableFencedMathBlocks"/> is enabled.
    /// Default: <c>math</c>, <c>latex</c>.
    /// </summary>
    public string[] FencedMathLanguages { get; set; } = new[] { "math", "latex" };
}
