using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Options controlling how Markdown is parsed and rendered to HTML for a WebView/browser host.
/// </summary>
public sealed class MarkdownRendererOptions {
    /// <summary>
    /// Markdown reader options used when parsing Markdown into OfficeIMO.Markdown's typed model.
    /// Defaults are biased for untrusted input (HTML disabled and file URLs blocked).
    /// </summary>
    public MarkdownReaderOptions ReaderOptions { get; set; } = new MarkdownReaderOptions {
        HtmlBlocks = false,
        InlineHtml = false,
        DisallowFileUrls = true,
        AllowDataUrls = false,
        AllowProtocolRelativeUrls = false,
        RestrictUrlSchemes = true,
        AllowedUrlSchemes = new[] { "http", "https", "mailto" }
    };

    /// <summary>
    /// HTML rendering options. These control the CSS preset and optional Prism support.
    /// </summary>
    public HtmlOptions HtmlOptions { get; set; } = new HtmlOptions {
        Kind = HtmlKind.Fragment,
        Style = HtmlStyle.GithubAuto,
        CssDelivery = CssDelivery.Inline,
        AssetMode = AssetMode.Online,
        BodyClass = "markdown-body",
        RawHtmlHandling = RawHtmlHandling.Strip,
        ExternalLinksTargetBlank = true,
        ExternalLinksRel = "noopener noreferrer nofollow ugc",
        ExternalLinksReferrerPolicy = "no-referrer",
        RestrictHttpLinksToBaseOrigin = true,
        RestrictHttpImagesToBaseOrigin = true,
        BlockExternalHttpImages = true,
        ImagesLoadingLazy = true,
        ImagesDecodingAsync = true,
        ImagesReferrerPolicy = "no-referrer",
        Prism = new PrismOptions { Enabled = true, Theme = PrismTheme.GithubAuto }
    };

    /// <summary>
    /// Optional base href inserted into the HTML update payload as a &lt;base&gt; tag.
    /// The incremental updater moves it to &lt;head&gt; so relative links/images resolve.
    /// </summary>
    public string? BaseHref { get; set; }

    /// <summary>
    /// When true, normalizes escaped newlines ("\\n"/"\\r\\n") into real newlines before parsing.
    /// Useful when Markdown arrives as a JSON string literal.
    /// Default: true.
    /// </summary>
    public bool NormalizeEscapedNewlines { get; set; } = true;

    /// <summary>Mermaid support options.</summary>
    public MermaidOptions Mermaid { get; } = new MermaidOptions();

    /// <summary>Chart.js support options.</summary>
    public ChartOptions Chart { get; } = new ChartOptions();

    /// <summary>Math (KaTeX) support options.</summary>
    public MathOptions Math { get; } = new MathOptions();

    /// <summary>
    /// Optional post-processors applied to the HTML fragment produced by <see cref="MarkdownRenderer.RenderBodyHtml"/>.
    /// These run after built-in conversions (Mermaid/Chart/Math) and after BaseHref injection.
    /// Default: none.
    /// </summary>
    public List<MarkdownHtmlPostProcessor> HtmlPostProcessors { get; } = new List<MarkdownHtmlPostProcessor>();

    /// <summary>
    /// Optional Content-Security-Policy meta tag value for the shell document (inserted as http-equiv).
    /// Leave unset unless your host wants to enforce a specific policy.
    /// Default: null.
    /// </summary>
    public string? ContentSecurityPolicy { get; set; }

    /// <summary>
    /// Optional additional CSS appended to the shell document after built-in styles and assets.
    /// Use this to theme/override the default styles without replacing the full shell HTML.
    /// Default: null.
    /// </summary>
    public string? ShellCss { get; set; }
}
