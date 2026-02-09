namespace OfficeIMO.Markdown;

/// <summary>
/// Options controlling HTML rendering style and asset delivery.
/// </summary>
public sealed class HtmlOptions {
    /// <summary>Fragment vs full document. Default: <see cref="HtmlKind.Document"/> when used with <see cref="MarkdownDoc.ToHtmlDocument"/>.</summary>
    public HtmlKind Kind { get; set; } = HtmlKind.Fragment;
    /// <summary>Built-in style preset. Default: <see cref="HtmlStyle.Clean"/>.</summary>
    public HtmlStyle Style { get; set; } = HtmlStyle.Clean;
    /// <summary>How to deliver CSS. Default: <see cref="CssDelivery.Inline"/>.</summary>
    public CssDelivery CssDelivery { get; set; } = CssDelivery.Inline;
    /// <summary>Connectivity mode for external assets (CDNs). Default: <see cref="AssetMode.Online"/>.</summary>
    public AssetMode AssetMode { get; set; } = AssetMode.Online;
    /// <summary>Optional explicit CSS URL for <see cref="CssDelivery.LinkHref"/>.</summary>
    public string? CssHref { get; set; }
    /// <summary>Additional CSS URLs to include (link or inline depending on <see cref="AssetMode"/>).</summary>
    public List<string> AdditionalCssHrefs { get; } = new();
    /// <summary>Additional JS URLs to include (script src or inline depending on <see cref="AssetMode"/>).</summary>
    public List<string> AdditionalJsHrefs { get; } = new();
    /// <summary>Page title for full document rendering. Default: "Document".</summary>
    public string Title { get; set; } = "Document";
    /// <summary>Wrap content in &lt;article&gt; with this CSS class. Set to null to avoid wrapper. Default: "markdown-body".</summary>
    public string? BodyClass { get; set; } = "markdown-body";
    /// <summary>Include per-heading anchor links (e.g. '#') in HTML rendering. Default: false.</summary>
    public bool IncludeAnchorLinks { get; set; } = false;
    /// <summary>When true, show a small anchor icon next to headings (on hover by default).</summary>
    public bool ShowAnchorIcons { get; set; } = false;
    /// <summary>Glyph or text for the anchor icon (e.g., "🔗", "¶"). Default: "🔗".</summary>
    public string AnchorIcon { get; set; } = "🔗";
    /// <summary>When true, clicking the anchor icon copies a deep link to the clipboard.</summary>
    public bool CopyHeadingLinkOnClick { get; set; } = false;
    /// <summary>Render small "Back to top" links for headings at or below the given level. 1=H1, 2=H2, etc. Set to false to disable.</summary>
    public bool BackToTopLinks { get; set; } = false;
    /// <summary>Heading level threshold for BackToTopLinks. Default: 2 (H2+).</summary>
    public int BackToTopMinLevel { get; set; } = 2;
    /// <summary>Text for the back-to-top link.</summary>
    public string BackToTopText { get; set; } = "Back to top";
    /// <summary>When true, writes a small theme toggle control if <see cref="Style"/> supports it. Default: false.</summary>
    public bool ThemeToggle { get; set; } = false;
    /// <summary>Emit tags vs manifest-only. Default: <see cref="AssetEmitMode.Emit"/>.</summary>
    public AssetEmitMode EmitMode { get; set; } = AssetEmitMode.Emit;
    /// <summary>Optional Prism highlighting configuration.</summary>
    public PrismOptions? Prism { get; set; }
    /// <summary>Prefix selectors in emitted CSS with this scope selector to avoid collisions. Default: "article.markdown-body".</summary>
    public string? CssScopeSelector { get; set; } = "article.markdown-body";

    /// <summary>
    /// Controls how raw HTML blocks are emitted. Default: <see cref="RawHtmlHandling.Allow"/>.
    /// For untrusted chat scenarios, prefer <see cref="RawHtmlHandling.Strip"/> or <see cref="RawHtmlHandling.Escape"/>.
    /// </summary>
    public RawHtmlHandling RawHtmlHandling { get; set; } = RawHtmlHandling.Allow;

    /// <summary>
    /// When true, external HTTP(S) links are rendered with <c>target="_blank"</c>.
    /// Default: false.
    /// </summary>
    public bool ExternalLinksTargetBlank { get; set; } = false;

    /// <summary>
    /// Optional <c>rel</c> attribute value to apply to external HTTP(S) links.
    /// Common safe value: <c>noopener noreferrer</c>.
    /// Default: empty (no rel attribute added).
    /// </summary>
    public string ExternalLinksRel { get; set; } = string.Empty;

    /// <summary>
    /// Optional <c>referrerpolicy</c> value to apply to external HTTP(S) links.
    /// Common privacy value: <c>no-referrer</c>.
    /// Default: empty (no referrerpolicy attribute added).
    /// </summary>
    public string ExternalLinksReferrerPolicy { get; set; } = string.Empty;

    // The following are used internally by the renderer; not part of the public API surface.
    internal string? ExternalCssOutputPath { get; set; }
    internal string? _externalCssContentToWrite { get; set; }

    /// <summary>Optional theme color overrides for links, headings, and TOC.</summary>
    public ThemeColors Theme { get; set; } = new ThemeColors();

    // TOC injection (used by higher-level pipelines like Word→Markdown→HTML)
    /// <summary>
    /// When true, injects a Table of Contents at the top of the document before rendering HTML.
    /// This is applied by host pipelines that have access to the MarkdownDoc model.
    /// </summary>
    public bool InjectTocAtTop { get; set; } = false;
    /// <summary>Title for the injected TOC. Default: "Contents".</summary>
    public string InjectTocTitle { get; set; } = "Contents";
    /// <summary>Minimum heading level to include in injected TOC. Default: 1.</summary>
    public int InjectTocMinLevel { get; set; } = 1;
    /// <summary>Maximum heading level to include in injected TOC. Default: 3.</summary>
    public int InjectTocMaxLevel { get; set; } = 3;
    /// <summary>Whether the injected TOC should be ordered (true) or unordered (false). Default: false.</summary>
    public bool InjectTocOrdered { get; set; } = false;
    /// <summary>Heading level used for the TOC title. Default: 2.</summary>
    public int InjectTocTitleLevel { get; set; } = 2;
}
