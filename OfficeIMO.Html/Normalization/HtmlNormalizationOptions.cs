namespace OfficeIMO.Html;

/// <summary>
/// Options controlling OfficeIMO normalized HTML output.
/// </summary>
public sealed class HtmlNormalizationOptions {
    /// <summary>Optional base URI used to resolve URL-bearing attributes.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>URL policy applied before URL-bearing attributes are emitted.</summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>When true, only body contents are emitted instead of a full document shell.</summary>
    public bool UseBodyContentsOnly { get; set; } = true;

    /// <summary>When true, comments are preserved. The default keeps normalized output document-focused.</summary>
    public bool PreserveComments { get; set; }

    /// <summary>When true, CSS style elements are preserved as normalized text content.</summary>
    public bool PreserveStyleElements { get; set; } = true;

    /// <summary>When true, inline event-handler attributes are removed from normalized output.</summary>
    public bool RemoveEventHandlerAttributes { get; set; } = true;

    /// <summary>When true, text nodes collapse repeated whitespace to one space.</summary>
    public bool CollapseTextWhitespace { get; set; } = true;
}
