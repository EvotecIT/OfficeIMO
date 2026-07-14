namespace OfficeIMO.Pdf;

internal sealed class FlowBlock : IPdfBlock {
    private readonly IReadOnlyList<IPdfBlock>? _blocks;
    private readonly Func<PdfFlowContext, IReadOnlyList<IPdfBlock>>? _factory;

    public FlowBlock(IEnumerable<IPdfBlock> blocks, PdfFlowOptions? options, PdfLayoutPositionCapture? capture) {
        Guard.NotNull(blocks, nameof(blocks));
        _blocks = blocks.ToArray();
        Options = options?.Clone() ?? new PdfFlowOptions();
        Capture = capture;
    }

    public FlowBlock(Func<PdfFlowContext, IReadOnlyList<IPdfBlock>> factory, PdfFlowOptions? options, PdfLayoutPositionCapture? capture) {
        Guard.NotNull(factory, nameof(factory));
        _factory = factory;
        Options = options?.Clone() ?? new PdfFlowOptions();
        Capture = capture;
    }

    public PdfFlowOptions Options { get; }
    public PdfLayoutPositionCapture? Capture { get; }
    public bool IsReplayable => _factory != null;
    public IReadOnlyList<IPdfBlock>? StaticBlocks => _blocks;

    public IReadOnlyList<IPdfBlock> Materialize(PdfFlowContext context) {
        return _factory?.Invoke(context) ?? _blocks!;
    }
}
