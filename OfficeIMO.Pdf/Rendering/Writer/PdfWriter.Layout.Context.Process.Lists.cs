using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderBulletListFlowBlock(BulletListBlock bl, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfListStyle? listStyle = ResolveListStyle(bl, currentOpts);
            double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
            PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
            const string bulletGlyph = "•";
            double estimatedBulletWidth = bl.RichItems.Count == 0
                ? EstimateSimpleTextWidthForOptions(bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts)
                : bl.RichItems.Max(item => EstimateSimpleTextWidthForOptions(item.Marker ?? bulletGlyph, markerFont, markerNamedFont, markerSize, currentOpts));
            double bulletWidth = GetListMarkerWidth(listStyle, estimatedBulletWidth);
            double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double indent = bulletWidth + markerGap;
            double listLeftIndent = listStyle?.LeftIndent ?? 0D;
            double rawTextWidth = width - listLeftIndent - indent;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
            double alignmentWidth = Math.Max(0, rawTextWidth);
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(bl.RichItems.Count);
            for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(bl.RichItems[itemIndex], availableWidth, baseFont, size, leading, currentOpts));
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
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size);
                double keepHeight = listHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            int? listStructureElementIndex = null;
            LayoutResult.Page? listStructurePage = null;
            for (int itemIndex = 0; itemIndex < bl.RichItems.Count; itemIndex++) {
                var item = bl.RichItems[itemIndex];
                string marker = item.Marker ?? bulletGlyph;
                var layout = wrappedItems[itemIndex];
                double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0], currentOpts) : 0;
                double firstLineDx = 0;
                if (bl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                else if (bl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                double spacingAfter = itemIndex == bl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                PdfColor? listColor = bl.Color ?? listStyle?.Color;
                RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, markerFont, markerNamedFont, markerSize, listStyle?.MarkerColor ?? listColor, currentOpts.MarginLeft + listLeftIndent + firstLineDx, bulletWidth, GetBulletMarkerAlign(listStyle), currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, bl.Align, listColor, size, leading, spacingBefore, spacingAfter, item.BookmarkName, ref listStructureElementIndex, ref listStructurePage);
            }
        }

        private void RenderNumberedListFlowBlock(NumberedListBlock nl, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfListStyle? listStyle = ResolveListStyle(nl, currentOpts);
            double size = GetListFontSize(listStyle, currentOpts.DefaultFontSize);
            double markerSize = GetListMarkerFontSize(listStyle, size);
            double leading = Math.Max(GetListLeading(listStyle, size), GetListLeading(listStyle, markerSize));
            var baseFont = ChooseNormal(currentOpts.DefaultFont);
            PdfStandardFont markerFont = GetListMarkerFont(listStyle, currentOpts.DefaultFont);
            PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(listStyle, currentOpts);
            int lastNumber = nl.StartNumber + Math.Max(0, nl.RichItems.Count - 1);
            string widestMarker = lastNumber.ToString(CultureInfo.InvariantCulture) + ".";
            double estimatedMarkerWidth = nl.RichItems.Count == 0
                ? EstimateSimpleTextWidthForOptions(widestMarker, markerFont, markerNamedFont, markerSize, currentOpts)
                : nl.RichItems
                    .Select((item, itemIndex) => item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + "."))
                    .Max(marker => EstimateSimpleTextWidthForOptions(marker, markerFont, markerNamedFont, markerSize, currentOpts));
            double markerWidth = GetListMarkerWidth(listStyle, estimatedMarkerWidth);
            double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
            double markerGap = GetListMarkerGap(listStyle, spaceAdvance);
            double indent = markerWidth + markerGap;
            double rawTextWidth = width - (listStyle?.LeftIndent ?? 0D) - indent;
            double availableWidth = Math.Max(rawTextWidth, EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
            double alignmentWidth = Math.Max(0, rawTextWidth);
            double itemSpacing = GetListItemSpacing(listStyle, leading);
            double listLeftIndent = listStyle?.LeftIndent ?? 0D;
            var wrappedItems = new System.Collections.Generic.List<TableCellTextLayout>(nl.RichItems.Count);
            for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                wrappedItems.Add(CreateListItemTextLayout(nl.RichItems[itemIndex], availableWidth, baseFont, size, leading, currentOpts));
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
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, size);
                double keepHeight = listHeight + nextHeight;
                double availableHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
                if (nextHeight > 0.001 && keepHeight <= availableHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    listSpacingBefore = 0D;
                    listHeight = MeasureListKeepTogetherHeight(wrappedItems, leading, listSpacingBefore, itemSpacing, listSpacingAfter);
                }
            }

            int? listStructureElementIndex = null;
            LayoutResult.Page? listStructurePage = null;
            for (int itemIndex = 0; itemIndex < nl.RichItems.Count; itemIndex++) {
                var item = nl.RichItems[itemIndex];
                string marker = item.Marker ?? ((nl.StartNumber + itemIndex).ToString(CultureInfo.InvariantCulture) + ".");
                var layout = wrappedItems[itemIndex];
                double firstLineWidth = layout.Lines.Count > 0 ? MeasureRichLineWidth(layout.Lines[0], currentOpts) : 0;
                double firstLineDx = 0;
                if (nl.Align == PdfAlign.Center) firstLineDx = Math.Max(0, (alignmentWidth - firstLineWidth) / 2);
                else if (nl.Align == PdfAlign.Right) firstLineDx = Math.Max(0, alignmentWidth - firstLineWidth);

                double spacingBefore = itemIndex == 0 ? listSpacingBefore : 0D;
                double spacingAfter = itemIndex == nl.RichItems.Count - 1 ? listSpacingAfter : itemSpacing;
                PdfColor? listColor = nl.Color ?? listStyle?.Color;
                RenderListItem(item.Runs, layout.Lines, layout.LineHeights, marker, markerFont, markerNamedFont, markerSize, listStyle?.MarkerColor ?? listColor, currentOpts.MarginLeft + listLeftIndent + firstLineDx, markerWidth, GetNumberedMarkerAlign(listStyle), currentOpts.MarginLeft + listLeftIndent + indent, alignmentWidth, nl.Align, listColor, size, leading, spacingBefore, spacingAfter, item.BookmarkName, ref listStructureElementIndex, ref listStructurePage);
            }
        }

    }
}
