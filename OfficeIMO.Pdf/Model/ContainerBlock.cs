namespace OfficeIMO.Pdf;

internal sealed class ContainerBlock : IPdfBlock {
    public ContainerBlock(IEnumerable<IPdfBlock> blocks, PdfPanelStyle? style, bool useDefaultPanelStyle = false) {
        Guard.NotNull(blocks, nameof(blocks));
        Blocks = blocks.ToList().AsReadOnly();
        Style = style?.Clone() ?? new PdfPanelStyle();
        UseDefaultPanelStyle = useDefaultPanelStyle && style == null;
    }

    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PdfPanelStyle Style { get; }
    public bool UseDefaultPanelStyle { get; }
}
