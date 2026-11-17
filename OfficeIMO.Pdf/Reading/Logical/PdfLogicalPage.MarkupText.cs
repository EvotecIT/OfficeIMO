namespace OfficeIMO.Pdf;

public sealed partial class PdfLogicalPage {
    private readonly PdfReadCache<IReadOnlyList<PdfPageInteractionRegion>> _markupTextRegions = new();
    // Retained only for pages that carry text-markup annotations; the text layer is rebuilt on first use.
    private PdfReadPage? _markupTextSource;

    /// <summary>
    /// Returns the page text covered by a Highlight, Underline, StrikeOut or Squiggly annotation from this page's
    /// <see cref="Annotations"/>. Text elements whose centers lie inside the annotation's /QuadPoints (or /Rect when it has
    /// no usable quad points) are joined in extraction order with the same word and line separators as
    /// <see cref="PdfPageInteractionMap.GetSelectedText"/>. Invisible OCR text layers count as covered text; runs whose
    /// font encoding cannot be mapped to Unicode are left out instead of failing.
    /// </summary>
    /// <returns>The covered text, or null when the annotation is not a text-markup annotation of this page or covers no text.</returns>
    public string? GetMarkupText(PdfAnnotation annotation, CancellationToken cancellationToken = default) {
        Guard.NotNull(annotation, nameof(annotation));
        cancellationToken.ThrowIfCancellationRequested();
        if (!annotation.IsTextMarkup || !IsOwnAnnotation(annotation)) return null;
        PdfReadPage? page = Volatile.Read(ref _markupTextSource);
        if (page is null) return null;
        IReadOnlyList<PdfPageInteractionRegion> regions = _markupTextRegions.GetOrCreate(page,
            static (source, token) => PdfPageInteractionMap.CreateMarkupTextRegions(source, token), cancellationToken);
        string text = PdfPageInteractionMap.GetMarkupText(page, regions, annotation, cancellationToken);
        return text.Length == 0 ? null : text;
    }

    // Page numbers repeat across documents, so only an annotation object read with this page is accepted.
    private bool IsOwnAnnotation(PdfAnnotation annotation) {
        for (int i = 0; i < Annotations.Count; i++) {
            if (ReferenceEquals(Annotations[i], annotation)) return true;
        }
        return false;
    }
}
