using System.Collections.Generic;

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
    /// <summary>Include per-heading anchor links (e.g. '#') in HTML rendering. Default: true.</summary>
    public bool IncludeAnchorLinks { get; set; } = true;
    /// <summary>When true, writes a small theme toggle control if <see cref="Style"/> supports it. Default: false.</summary>
    public bool ThemeToggle { get; set; } = false;
    /// <summary>Emit tags vs manifest-only. Default: <see cref="AssetEmitMode.Emit"/>.</summary>
    public AssetEmitMode EmitMode { get; set; } = AssetEmitMode.Emit;
    /// <summary>Optional Prism highlighting configuration.</summary>
    public PrismOptions? Prism { get; set; }
    /// <summary>Prefix selectors in emitted CSS with this scope selector to avoid collisions. Default: "article.markdown-body".</summary>
    public string CssScopeSelector { get; set; } = "article.markdown-body";

    // The following are used internally by the renderer; not part of the public API surface.
    internal string? ExternalCssOutputPath { get; set; }
    internal string? _externalCssContentToWrite { get; set; }
}

/// <summary>Parts of HTML output for advanced embedding.</summary>
 
