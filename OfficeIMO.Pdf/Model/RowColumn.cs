namespace OfficeIMO.Pdf;

internal sealed class RowColumn {
    private readonly System.Collections.Generic.List<IPdfBlock> _blocks = new();
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock> _blocksView;

    public PdfColumnWidth Width { get; }
    public System.Collections.Generic.IReadOnlyList<IPdfBlock> Blocks => _blocksView;

    public RowColumn(PdfColumnWidth width) {
        Width = width;
        _blocksView = new System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock>(_blocks);
    }

    internal void AddBlock(IPdfBlock block) {
        Guard.NotNull(block, nameof(block));
        _blocks.Add(block);
    }

}

