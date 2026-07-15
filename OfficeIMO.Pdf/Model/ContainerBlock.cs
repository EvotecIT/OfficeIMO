namespace OfficeIMO.Pdf;

internal sealed class ContainerBlock : IPdfBlock {
    public ContainerBlock(IEnumerable<IPdfBlock> blocks, PanelStyle? style) {
        Guard.NotNull(blocks, nameof(blocks));
        Blocks = blocks.ToList().AsReadOnly();
        Style = style?.Clone() ?? new PanelStyle();
    }

    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PanelStyle Style { get; }
}
