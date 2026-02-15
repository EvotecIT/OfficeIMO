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

    /// <summary>
    /// When true, joins short hard-wrapped bold labels (for example, "**Status\nOK**") into a single bold span.
    /// This helps chat-style outputs that wrap short headers mid-token.
    /// Default: false.
    /// </summary>
    public bool NormalizeSoftWrappedStrongSpans { get; set; } = false;

    /// <summary>
    /// When true, compacts inline code spans containing line breaks into a single line.
    /// This preserves strict parser compatibility for malformed model outputs.
    /// Default: false.
    /// </summary>
    public bool NormalizeInlineCodeSpanLineBreaks { get; set; } = false;

    /// <summary>
    /// Optional markdown pre-processors applied before parsing.
    /// These run after escaped newline normalization and after built-in text normalization.
    /// Default: none.
    /// </summary>
    public List<MarkdownTextPreProcessor> MarkdownPreProcessors { get; } = new List<MarkdownTextPreProcessor>();

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

    /// <summary>
    /// When true, the shell will add "Copy" buttons to fenced code blocks (client-side).
    /// Default: false.
    /// </summary>
    public bool EnableCodeCopyButtons { get; set; } = false;

    /// <summary>
    /// When true, the shell will add table copy actions (TSV/CSV) above rendered tables (client-side).
    /// Default: false.
    /// </summary>
    public bool EnableTableCopyButtons { get; set; } = false;

    /// <summary>
    /// Optional guardrail limit for Markdown input size (character count). When exceeded, <see cref="MarkdownOverflowHandling"/> applies.
    /// Default: null (no limit).
    /// </summary>
    public int? MaxMarkdownChars { get; set; }

    /// <summary>
    /// Optional guardrail limit for the rendered HTML payload size (UTF-8 bytes). When exceeded, <see cref="BodyHtmlOverflowHandling"/> applies.
    /// Default: null (no limit).
    /// </summary>
    public int? MaxBodyHtmlBytes { get; set; }

    /// <summary>
    /// Behavior when <see cref="MaxMarkdownChars"/> is exceeded. Default: <see cref="OverflowHandling.Truncate"/>.
    /// </summary>
    public OverflowHandling MarkdownOverflowHandling { get; set; } = OverflowHandling.Truncate;

    /// <summary>
    /// Behavior when <see cref="MaxBodyHtmlBytes"/> is exceeded. Default: <see cref="OverflowHandling.RenderError"/>.
    /// </summary>
    public OverflowHandling BodyHtmlOverflowHandling { get; set; } = OverflowHandling.RenderError;
}
