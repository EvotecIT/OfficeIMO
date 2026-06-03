namespace OfficeIMO.Pdf;

internal sealed class PdfGeneratedImageAccessibilityEvidence {
    public PdfGeneratedImageAccessibilityEvidence(bool hasAlternativeText, bool isDecorativeArtifact) {
        HasAlternativeText = hasAlternativeText;
        IsDecorativeArtifact = isDecorativeArtifact;
    }

    public bool HasAlternativeText { get; }

    public bool IsDecorativeArtifact { get; }
}
