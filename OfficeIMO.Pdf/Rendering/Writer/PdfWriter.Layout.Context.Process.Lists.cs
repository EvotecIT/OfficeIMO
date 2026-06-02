using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderBulletListFlowBlock(BulletListBlock bl, IPdfBlock? nextBlock) {
            PdfListStyle? listStyle = ResolveListStyle(bl, currentOpts);
            double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
            double leading = GetListLeading(listStyle, size);
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            const string bulletGlyph = "•";
            double bulletWidth = bl.RichItems.Count == 0
                ? EstimateSimpleTextWidth(bulletGlyph, baseFont, size)
                : bl.RichItems.Max(item => EstimateSimpleTextWidth(item.Marker ?? bulletGlyph, baseFont, size));
            double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double indent = bulletWidth + markerGap;
            double listLeftIndent = listStyle?.LeftIndent ?? 0D;
            double rawTextWidth = width - listLeftIndent - indent;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
            double alignmentWidth = Math.Max(0, rawTextWidth);
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(bl.RichItems.Count);
            for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(bl.RichItems[itemIndex], availableWidth, baseFont, size, leading));
            }

            double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
            double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
            double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
            if (listStyle?.KeepTogether == true) {
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (listHeight > availableHeight + 0.001) {
                    throw new ArgumentException("List height exceeds the available page content height.");
                }

                if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                double keepHeight = listHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                var item = bl.RichItems[itemIndex];
                string marker = item.Marker ?? bulletGlyph;
                var layout = wrappedItems[itemIndex];
                double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                double firstLineDx = 0;
                if (bl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                else if (bl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                double spacingAfter = itemIndex == bl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, currentOpts.MarginLeft + listLeftIndent + firstLineDx, bulletWidth, PdfAlign.Left, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, bl.Align, bl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter, item.BookmarkName);
            }
        }

        private void RenderNumberedListFlowBlock(NumberedListBlock nl, IPdfBlock? nextBlock) {
            PdfListStyle? listStyle = ResolveListStyle(nl, currentOpts);
            double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
            double leading = GetListLeading(listStyle, size);
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            int lastNumber = nl.StartNumber + Math.Max(0, nl.RichItems.Count - 1);
            string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
            double markerWidth = nl.RichItems.Count == 0
                ? EstimateSimpleTextWidth(widestMarker, baseFont, size)
                : nl.RichItems
                    .Select((item, itemIndex) => item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                    .Max(marker => EstimateSimpleTextWidth(marker, baseFont, size));
            double spaceAdvance = EstimateSimpleTextWidth(" ", baseFont, size);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double indent = markerWidth + markerGap;
            double rawTextWidth = width - (listStyle?.LeftIndent ?? 0D) - indent;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidth("WW", baseFont, size));
            double alignmentWidth = Math.Max(0, rawTextWidth);
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            double listLeftIndent = listStyle?.LeftIndent ?? 0D;
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(nl.RichItems.Count);
            for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(nl.RichItems[itemIndex], availableWidth, baseFont, size, leading));
            }

            double listSpacingBefore = ResolveTopLevelSpacingBefore(listStyle?.SpacingBefore ?? 0D);
            double listSpacingAfter = listStyle?.GetSpacingAfter(itemSpacing) ?? itemSpacing;
            double listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
            if (listStyle?.KeepTogether == true) {
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (listHeight > availableHeight + 0.001) {
                    throw new ArgumentException("List height exceeds the available page content height.");
                }

                if (y < yStart - 0.001 && y - listHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            if (listStyle?.KeepWithNext == true && nextBlock != null && wrappedItems.Count > 0) {
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, size);
                double keepHeight = listHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                var item = nl.RichItems[itemIndex];
                string marker = item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".");
                var layout = wrappedItems[itemIndex];
                double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0]) : 0;
                double firstLineDx = 0;
                if (nl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                else if (nl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                double spacingAfter = itemIndex == nl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, currentOpts.MarginLeft + listLeftIndent + firstLineDx, markerWidth, PdfAlign.Right, currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, nl.Align, nl.Color ?? listStyle?.Color, size, leading, spacingBefore, spacingAfter, item.BookmarkName);
            }
        }

    }
}
