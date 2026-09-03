namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal IReadOnlyList<PdfTextSpan> GetHiddenOptionalContentTextSpans(bool includeArtifactText) {
        IReadOnlyList<PdfTextSpan> visible = GetTextSpans(includeArtifactText, default);
        IReadOnlyList<PdfTextSpan> includingHidden = GetTextSpans(
            includeArtifactText,
            default,
            includeHiddenOptionalContent: true);
        var visibleKeys = new HashSet<PdfContentOrderKey>(visible
            .Select(static span => span.ContentOrderKey)
            .OfType<PdfContentOrderKey>());
        return includingHidden
            .Where(span => span.ContentOrderKey is not null && !visibleKeys.Contains(span.ContentOrderKey))
            .ToArray();
    }
}
