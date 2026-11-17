namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    private readonly PdfReadCache<IReadOnlyList<PdfPageInteractionRegion>> _markupTextRegions = new();
    // Retained only for pages that carry text-markup annotations; the text layer is rebuilt on first use.
    private PdfReadPage? _markupTextSource;

    /// <summary>
    /// Returns the page text covered by a Highlight, Underline, StrikeOut or Squiggly annotation on this page.
    /// Text elements whose centers lie inside the annotation's /QuadPoints (or /Rect when it has none) are joined
    /// in extraction order with the same word and line separators as <see cref="PdfPageInteractionMap.GetSelectedText"/>.
    /// Invisible OCR text layers count as covered text.
    /// </summary>
    /// <returns>The covered text, or null when the annotation is not a text-markup annotation of this page or covers no text.</returns>
    public string? GetMarkupText(PdfAnnotation annotation, CancellationToken cancellationToken = default) {
        Guard.NotNull(annotation, nameof(annotation));
        cancellationToken.ThrowIfCancellationRequested();
        if (!annotation.IsTextMarkup || (annotation.PageNumber.HasValue && annotation.PageNumber.Value != PageNumber)) return null;
        PdfReadPage? page = Volatile.Read(ref _markupTextSource);
        if (page is null) return null;
        IReadOnlyList<PdfPageInteractionRegion> regions = _markupTextRegions.GetOrCreate(page,
            static (source, _) => PdfPageInteractionMap.CreateMarkupTextRegions(source), cancellationToken);
        string text = PdfPageInteractionMap.GetMarkupText(page, regions, annotation);
        return text.Length == 0 ? null : text;
    }
}
