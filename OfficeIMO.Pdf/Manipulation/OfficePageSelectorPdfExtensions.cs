namespace OfficeIMO.Pdf;

/// <summary>Maps a shared document selector to the absolute PDF page-selection contract.</summary>
public static class OfficePageSelectorPdfExtensions {
    /// <summary>Resolves the selector to caller-ordered, one-based absolute PDF pages.</summary>
    public static PdfPageSelection ResolveSelection(this OfficePageSelector selector, int pageCount) {
        ArgumentNullException.ThrowIfNull(selector);
        return PdfPageSelection.From(selector.Resolve(pageCount).ToArray());
    }
}
