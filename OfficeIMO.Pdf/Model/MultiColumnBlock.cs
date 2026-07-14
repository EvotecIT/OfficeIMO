namespace OfficeIMO.Pdf;

internal sealed class MultiColumnBlock : IPdfBlock {
    public MultiColumnBlock(IEnumerable<IPdfBlock> blocks, PdfMultiColumnOptions? options) {
        Guard.NotNull(blocks, nameof(blocks));
        Blocks = blocks.ToList().AsReadOnly();
        Options = options?.Clone() ?? new PdfMultiColumnOptions();
    }

    public IReadOnlyList<IPdfBlock> Blocks { get; }
    public PdfMultiColumnOptions Options { get; }
}

internal sealed class ColumnBreakBlock : IPdfBlock { }
