namespace OfficeIMO.Pdf;

internal sealed class SemanticBlock : IPdfBlock {
    public SemanticBlock(PdfSemanticRole role, IEnumerable<IPdfBlock> blocks, string? alternativeText) {
        Guard.NotNull(blocks, nameof(blocks));
        if ((int)role < (int)PdfSemanticRole.Part || (int)role > (int)PdfSemanticRole.Form) {
            throw new ArgumentOutOfRangeException(nameof(role));
        }
        if (role == PdfSemanticRole.Figure && string.IsNullOrWhiteSpace(alternativeText)) {
            throw new ArgumentException("Figure semantics require alternate text.", nameof(alternativeText));
        }
        if (alternativeText != null) {
            Guard.NotNullOrWhiteSpace(alternativeText, nameof(alternativeText));
        }

        Role = role;
        Blocks = Array.AsReadOnly(blocks.ToArray());
        AlternativeText = alternativeText?.Trim();
    }

    public PdfSemanticRole Role { get; }
    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public string? AlternativeText { get; }
}
