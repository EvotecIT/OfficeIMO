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

    /// <summary>Generated structure without a public authoring role, such as the TOC and TOCI of generated contents.</summary>
    internal SemanticBlock(string structureType, IEnumerable<IPdfBlock> blocks) {
        Guard.NotNull(blocks, nameof(blocks));
        Role = PdfSemanticRole.Division;
        StructureType = structureType;
        Blocks = Array.AsReadOnly(blocks.ToArray());
    }

    public PdfSemanticRole Role { get; }
    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public string? AlternativeText { get; }
    internal string? StructureType { get; }
}
