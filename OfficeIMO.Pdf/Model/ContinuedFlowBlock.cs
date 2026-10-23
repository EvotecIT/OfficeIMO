namespace OfficeIMO.Pdf;

/// <summary>Captured ordinary flow with first-page and continuation decorations, independent of reference admission.</summary>
internal sealed class ContinuedFlowBlock : IPdfBlock {
    internal ContinuedFlowBlock(IReadOnlyList<IPdfBlock> blocks, IReadOnlyList<IPdfBlock> separator,
        IReadOnlyList<IPdfBlock> continuationSeparator, IReadOnlyList<IPdfBlock> continuationNotice,
        bool reserveFinalNotice) {
        Blocks = blocks;
        Separator = separator;
        ContinuationSeparator = continuationSeparator;
        ContinuationNotice = continuationNotice;
        ReserveFinalNotice = reserveFinalNotice;
    }

    internal IReadOnlyList<IPdfBlock> Blocks { get; }
    internal IReadOnlyList<IPdfBlock> Separator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationSeparator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationNotice { get; }
    internal bool ReserveFinalNotice { get; }
}
