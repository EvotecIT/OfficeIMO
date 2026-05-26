namespace OfficeIMO.Pdf;

internal sealed class PageBlock : IPdfBlock {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock> _blocksView;

    public PdfOptions Options { get; }
    public System.Collections.Generic.IReadOnlyList<IPdfBlock> Blocks => _blocksView;

    public PageBlock(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        Options = options;
        _blocksView = new System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock>(_blocks);
    }

    internal void AddBlock(IPdfBlock block) {
        Guard.NotNull(block, nameof(block));
        _blocks.Add(block);
    }
}
