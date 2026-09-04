namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    internal PdfDocument Semantic(PdfSemanticRole role, Action<PdfContentBuilder> compose, string? alternativeText = null) {
        Guard.NotNull(compose, nameof(compose));
        AddBlock(new SemanticBlock(role, BuildFlowBlocks(compose), alternativeText));
        return this;
    }
}
