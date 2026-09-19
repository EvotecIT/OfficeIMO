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
            double listHeight = MeasurePreparedListHeight(prepared);
            if (prepared.Style?.KeepTogether == true) {
                double availableHeight = GetFullPageContentHeight();
                if (listHeight > availableHeight + 0.001D) {
                    throw new ArgumentException("List height exceeds the available page content height.");
                }

                if (y < GetCurrentFramePageStartY() - 0.001D && y - listHeight < currentOpts.MarginBottom) {
                    NewPage();
                    prepared.SpacingBefore = 0D;
                    listHeight = MeasurePreparedListHeight(prepared);
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
                double availableHeight = GetFullPageContentHeight();
                if (nextHeight > 0.001D &&
                    keepHeight <= availableHeight + 0.001D &&
                    y < GetCurrentFramePageStartY() - 0.001D &&
                    y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    prepared.SpacingBefore = 0D;
                }
            }

            int? listStructureElementIndex = null;
            LayoutResult.Page? listStructurePage = null;
            for (int itemIndex = 0; itemIndex < prepared.Items.Count; itemIndex++) {
                PreparedListItem preparedItem = prepared.Items[itemIndex];
                PdfListItem item = preparedItem.Item;
                TableCellTextLayout layout = preparedItem.TextLayout;
                double spacingBefore = itemIndex == 0 ? prepared.SpacingBefore : 0D;
                double spacingAfter = itemIndex == prepared.Items.Count - 1
                    ? prepared.SpacingAfter
                    : prepared.ItemSpacing;
                PdfColor? listColor = list.Color ?? prepared.Style?.Color;
                RenderListItem(
                    item.Runs,
                    layout.Lines,
                    layout.LineHeights,
                    preparedItem.Marker,
                    prepared.MarkerFont,
                    prepared.MarkerNamedFont,
                    prepared.MarkerSize,
                    prepared.Style?.MarkerColor ?? listColor,
                    currentOpts.MarginLeft + prepared.ListLeftIndent + preparedItem.FirstLineOffset,
                    prepared.MarkerWidth,
                    list.GetMarkerAlign(prepared.Style),
                    currentOpts.MarginLeft + prepared.ListLeftIndent + prepared.MarkerWidth + prepared.MarkerGap,
                    prepared.AlignmentWidth,
                    list.Align,
                    listColor,
                    prepared.Size,
                    prepared.Leading,
                    spacingBefore,
                    spacingAfter,
                    item.BookmarkName,
                    ref listStructureElementIndex,
                    ref listStructurePage);
            }
        }
    }
}
