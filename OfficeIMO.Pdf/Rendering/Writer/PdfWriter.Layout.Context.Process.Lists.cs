namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderListFlowBlock(
            PdfListBlock list,
            IPdfBlock? nextBlock,
            System.Collections.Generic.IList<IPdfBlock> blockList,
            int blockIndex) {
            PreparedListLayout prepared = PrepareListLayout(
                list,
                width,
                currentOpts.DefaultFontSize,
                topLevelSpacing: true);
            double preparedWidth = width;
            double listHeight = MeasurePreparedListHeight(prepared);
            void NewListStartFrame(bool forcePhysicalPage = false) {
                NewBlockFrame(forcePhysicalPage);
                if (Math.Abs(preparedWidth - width) > .001D) {
                    prepared = PrepareListLayout(list, width, currentOpts.DefaultFontSize, topLevelSpacing: true);
                    preparedWidth = width;
                }
                prepared.SpacingBefore = 0D;
                listHeight = MeasurePreparedListHeight(prepared);
            }
            if (prepared.Style?.KeepTogether == true) {
                while (true) {
                    ContainerClosingSpace closing = GetClosingContainerSpace();
                    double fitHeight = MeasurePreparedListAdmissionHeight(prepared, closing);
                    if (fitHeight > GetMaximumBlockContinuationHeight() + .001D) {
                        if (prepared.Style?.KeepTogetherOverflow is null or PdfKeepTogetherOverflowBehavior.Reject)
                            throw new ArgumentException("List height exceeds the available page content height.");
                        if (y < GetCurrentFramePageStartY() - .001D || IsPartialPhysicalColumnFrame())
                            NewListStartFrame(IsPartialPhysicalColumnFrame() ||
                                prepared.Style.KeepTogetherOverflow == PdfKeepTogetherOverflowBehavior.SplitAcrossPages);
                        break;
                    }
                    if (!ShouldAdvanceForBlockHeight(fitHeight)) break;
                    NewListStartFrame();
                }
            }

            if (prepared.Style?.KeepWithNext == true && nextBlock != null && prepared.Items.Count > 0) {
                double nextHeight = MeasureKeepWithNextChainHeight(
                    blockList,
                    blockIndex + 1,
                    currentOpts.MarginLeft,
                    width,
                    prepared.Size,
                    listHeight);
                double keepHeight = listHeight + nextHeight;
                double availableHeight = GetMaximumBlockContinuationHeight();
                while (ReservesWholeKeepGroup() && nextHeight > 0.001D &&
                    keepHeight <= availableHeight + 0.001D &&
                    ShouldAdvanceForBlockHeight(keepHeight)) {
                    NewBlockFrame();
                    if (Math.Abs(preparedWidth - width) > .001D) {
                        prepared = PrepareListLayout(list, width, currentOpts.DefaultFontSize, topLevelSpacing: true);
                        preparedWidth = width;
                    }
                    prepared.SpacingBefore = 0D;
                    listHeight = MeasurePreparedListHeight(prepared);
                    nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, prepared.Size, listHeight);
                    keepHeight = listHeight + nextHeight;
                    availableHeight = GetMaximumBlockContinuationHeight();
                }
            }

            int? listStructureElementIndex = null;
            LayoutResult.Page? listStructurePage = null;
            for (int itemIndex = 0; itemIndex < prepared.Items.Count; itemIndex++) {
                if (activeColumnFlow != null && Math.Abs(preparedWidth - width) > 0.001D) {
                    prepared = PrepareListLayout(list, width, currentOpts.DefaultFontSize, topLevelSpacing: true);
                    preparedWidth = width;
                }
                RenderListItem(prepared, itemIndex, blockList, blockIndex,
                    ref listStructureElementIndex, ref listStructurePage);
            }
        }
    }
}
