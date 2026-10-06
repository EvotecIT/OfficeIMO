namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Kept blocks may use a later physical page when columns begin below the page top.</summary>
        private double GetMaximumBlockContinuationHeight() =>
            Math.Max(GetFullPageContentHeight(), GetMeasuredContinuationFrameHeight());

        /// <summary>Allows an unused partial column to be skipped, but never retries an impossible full frame.</summary>
        private bool ShouldAdvanceForBlockHeight(double height) =>
            height > y - currentOpts.MarginBottom + .001D &&
            (y < GetCurrentFramePageStartY() - .001D ||
             IsPartialPhysicalColumnFrame() && height <= GetMaximumBlockContinuationHeight() + .001D &&
             GetFullPageContentHeight() < GetMaximumBlockContinuationHeight() - .001D);

        private bool IsBalancedColumnFrame() => activeColumnFlow is { } scope &&
            scope.TargetHeight < scope.Top - scope.ParentOptions.MarginBottom - .001D;

        /// <summary>Distinguishes an actual partial physical page from a shortened balance target.</summary>
        private bool IsPartialPhysicalColumnFrame() => activeColumnFlow is { } scope &&
            GetCurrentFramePageStartY() - scope.ParentOptions.MarginBottom < GetMeasuredContinuationFrameHeight() - .001D;

        private bool ReservesWholeKeepGroup() => !IsBalancedColumnFrame() || IsPartialPhysicalColumnFrame();

        /// <summary>Preserves the unrendered block and its ancestor siblings when a block starts on a later page.</summary>
        private void NewBlockFrame() {
            if (activeBlockSequences.Count > 0) {
                BlockSequenceScope sequence = activeBlockSequences[activeBlockSequences.Count - 1];
                QueueColumnBalanceRemainder(sequence.Blocks[sequence.Index], sequence.Blocks, sequence.Index);
            }
            NewPage();
        }
    }
}
