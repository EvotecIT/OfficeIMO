namespace OfficeIMO.Html;

/// <summary>
/// Options controlling OfficeIMO normalized HTML output.
/// </summary>
public sealed class HtmlNormalizationOptions {
    /// <summary>Optional base URI used to resolve URL-bearing attributes.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>Optional base URI used only when resolving the document <c>base</c> element itself.</summary>
    public Uri? BaseElementBaseUri { get; set; }

    /// <summary>URL policy applied before URL-bearing attributes are emitted.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>
    /// Optional separate policy for images, stylesheets, fonts, and other non-hyperlink resources.
    /// When omitted, a resource policy is derived from <see cref="UrlPolicy"/>.
    /// </summary>
    public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }

    /// <summary>Shared source, DOM, stylesheet, and semantic-metadata limits applied before normalization.</summary>
    public HtmlConversionLimits Limits { get; set; } = HtmlConversionLimits.CreateUntrustedProfile();

    /// <summary>Optional maximum number of candidates normalized from one responsive source set.</summary>
    public int? MaxResponsiveImageCandidates { get; set; }

    /// <summary>When true, only body contents are emitted instead of a full document shell.</summary>
    public bool UseBodyContentsOnly { get; set; } = true;

    /// <summary>When true, comments are preserved. The default keeps normalized output document-focused.</summary>
    public bool PreserveComments { get; set; }

    /// <summary>
    /// When true, non-rendered elements such as <c>script</c> and <c>template</c> are retained as empty marker nodes.
    /// Target adapters use the markers to report skipped source content without receiving executable or hidden payload text.
    /// </summary>
    public bool PreserveSkippedElementMarkers { get; set; }

    /// <summary>When true, CSS style elements are preserved as normalized text content.</summary>
    public bool PreserveStyleElements { get; set; } = true;

    /// <summary>When true, inline event-handler attributes are removed from normalized output.</summary>
    public bool RemoveEventHandlerAttributes { get; set; } = true;

    /// <summary>When true, text nodes collapse repeated whitespace to one space.</summary>
    public bool CollapseTextWhitespace { get; set; } = true;
}
