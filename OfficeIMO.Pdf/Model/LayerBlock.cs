namespace OfficeIMO.Pdf;

internal sealed class PdfLayerDefinition {
    public PdfLayerDefinition(int id, string name, PdfLayerOptions options) {
        Id = id;
        Name = name;
        Options = options.Clone();
    }

    public int Id { get; }
    public string Name { get; }
    public PdfLayerOptions Options { get; }
    public string ResourceName => "OC" + Id.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

internal sealed class LayerBlock : IPdfBlock {
    private readonly System.Collections.ObjectModel.ReadOnlyCollection<IPdfBlock> _blocks;

    public LayerBlock(PdfLayerDefinition definition, IEnumerable<IPdfBlock> blocks) {
        Guard.NotNull(definition, nameof(definition));
        Guard.NotNull(blocks, nameof(blocks));
        Definition = definition;
        _blocks = blocks.ToList().AsReadOnly();
    }

    public PdfLayerDefinition Definition { get; }
    public IReadOnlyList<IPdfBlock> Blocks => _blocks;
}
