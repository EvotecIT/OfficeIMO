namespace OfficeIMO.Pdf;

internal sealed class PdfGeneratedDrawingAccessibilityEvidence {
    public PdfGeneratedDrawingAccessibilityEvidence(bool hasAlternativeText, bool isDecorativeArtifact) {
        HasAlternativeText = hasAlternativeText;
        IsDecorativeArtifact = isDecorativeArtifact;
    }

    public bool HasAlternativeText { get; }

    public bool IsDecorativeArtifact { get; }
}
