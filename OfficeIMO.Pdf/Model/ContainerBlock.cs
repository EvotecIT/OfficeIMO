namespace OfficeIMO.Pdf;

internal sealed class ContainerBlock : IPdfBlock {
    public ContainerBlock(IEnumerable<IPdfBlock> blocks, PdfPanelStyle? style) {
        Guard.NotNull(blocks, nameof(blocks));
        Blocks = blocks.ToList().AsReadOnly();
        Style = style?.Clone() ?? new PdfPanelStyle();
    }

    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PdfPanelStyle Style { get; }
}
