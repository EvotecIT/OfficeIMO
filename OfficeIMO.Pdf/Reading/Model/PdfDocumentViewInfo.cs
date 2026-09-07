namespace OfficeIMO.Pdf;

/// <summary>
/// Authenticated page geometry and capabilities for a document viewer. Restricted logical content is omitted.
/// </summary>
public sealed class PdfDocumentViewInfo {
    internal PdfDocumentViewInfo(IReadOnlyList<PdfPageInfo> pages, PdfDocumentSecurityInfo security,
        bool canExtractText, PdfDocumentInfo? logicalContent) {
        Pages = pages; Security = security; CanExtractText = canExtractText; LogicalContent = logicalContent;
    }

    /// <summary>Page geometry in document order. Annotation and widget objects are omitted when content extraction is restricted.</summary>
    public IReadOnlyList<PdfPageInfo> Pages { get; }
    /// <summary>Number of viewable pages.</summary>
    public int PageCount => Pages.Count;
    /// <summary>Authenticated security and permission information.</summary>
    public PdfDocumentSecurityInfo Security { get; }
    /// <summary>Whether the current authorization permits text extraction, including accessibility text.</summary>
    public bool CanExtractText { get; }
    /// <summary>Whether the current authorization permits logical content extraction.</summary>
    public bool CanExtractContent => LogicalContent is not null;
    /// <summary>Full inspection only when content extraction is authorized; otherwise null.</summary>
    public PdfDocumentInfo? LogicalContent { get; }
}
