namespace OfficeIMO.Pdf;

/// <summary>Captured ordinary flow with first-page and continuation decorations, independent of reference admission.</summary>
internal sealed class ContinuedFlowBlock : IPdfBlock {
    internal ContinuedFlowBlock(IReadOnlyList<IPdfBlock> blocks, IReadOnlyList<IPdfBlock> separator,
        IReadOnlyList<IPdfBlock> continuationSeparator, IReadOnlyList<IPdfBlock> continuationNotice,
        bool reserveFinalNotice, bool allowContinuationSeparatorInBottomMargin = false) {
        Blocks = blocks;
        Separator = separator;
        ContinuationSeparator = continuationSeparator;
        ContinuationNotice = continuationNotice;
        ReserveFinalNotice = reserveFinalNotice;
        AllowContinuationSeparatorInBottomMargin = allowContinuationSeparatorInBottomMargin;
    }

    internal IReadOnlyList<IPdfBlock> Blocks { get; }
    internal IReadOnlyList<IPdfBlock> Separator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationSeparator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationNotice { get; }
    internal bool ReserveFinalNotice { get; }
    /// <summary>Allows a continued separator to occupy unused bottom margin when the page has no notice, footnotes or selected footer.</summary>
    internal bool AllowContinuationSeparatorInBottomMargin { get; }
}
