namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class PreparedListLayout {
            internal PreparedListLayout(
                PdfListBlock block,
                PdfListStyle? style,
                PdfStandardFont markerFont,
                PdfNamedFontFace? markerNamedFont,
                double markerSize,
                double markerWidth,
                double markerGap,
                double listLeftIndent,
                double alignmentWidth,
                double size,
                double leading,
                double itemSpacing,
                double spacingBefore,
                double spacingAfter,
                System.Collections.Generic.List<PreparedListItem> items) {
                Block = block;
                Style = style;
                MarkerFont = markerFont;
                MarkerNamedFont = markerNamedFont;
                MarkerSize = markerSize;
                MarkerWidth = markerWidth;
                MarkerGap = markerGap;
                ListLeftIndent = listLeftIndent;
                AlignmentWidth = alignmentWidth;
                Size = size;
                Leading = leading;
                ItemSpacing = itemSpacing;
                SpacingBefore = spacingBefore;
                SpacingAfter = spacingAfter;
                Items = items;
            }

            internal PdfListBlock Block { get; }
            internal PdfListStyle? Style { get; }
            internal PdfStandardFont MarkerFont { get; }
            internal PdfNamedFontFace? MarkerNamedFont { get; }
            internal double MarkerSize { get; }
            internal double MarkerWidth { get; }
            internal double MarkerGap { get; }
            internal double ListLeftIndent { get; }
            internal double AlignmentWidth { get; }
            internal double Size { get; }
            internal double Leading { get; }
            internal double ItemSpacing { get; }
            internal double SpacingBefore { get; set; }
            internal double SpacingAfter { get; }
            internal System.Collections.Generic.List<PreparedListItem> Items { get; }
        }

        private sealed class PreparedListItem {
            internal PreparedListItem(
                PdfListItem item,
                string marker,
                TableCellTextLayout textLayout,
                double firstLineOffset) {
                Item = item;
                Marker = marker;
                TextLayout = textLayout;
                FirstLineOffset = firstLineOffset;
            }

            internal PdfListItem Item { get; }
            internal string Marker { get; }
            internal TableCellTextLayout TextLayout { get; }
            internal double FirstLineOffset { get; }
        }

        private PreparedListLayout PrepareListLayout(
            PdfListBlock block,
            double frameWidth,
            double defaultFontSize,
            bool topLevelSpacing) {
            PdfListStyle? style = ResolveListStyle(block, currentOpts);
            double size = GetListFontSize(style, defaultFontSize);
            double markerSize = GetListMarkerFontSize(style, size);
            double leading = System.Math.Max(GetListLeading(style, size), GetListLeading(style, markerSize));
            PdfStandardFont baseFont = ChooseNormal(currentOpts.DefaultFont);
            PdfStandardFont markerFont = GetListMarkerFont(style, currentOpts.DefaultFont);
            PdfNamedFontFace? markerNamedFont = GetListMarkerNamedFont(style, currentOpts);

            var markers = new string[block.RichItems.Count];
            double estimatedMarkerWidth = EstimateSimpleTextWidthForOptions(
                block.GetWidestDefaultMarker(),
                markerFont,
                markerNamedFont,
                markerSize,
                currentOpts);
            for (int itemIndex = 0; itemIndex < markers.Length; itemIndex++) {
                string marker = block.GetMarker(itemIndex);
                markers[itemIndex] = marker;
                estimatedMarkerWidth = System.Math.Max(
                    estimatedMarkerWidth,
                    EstimateSimpleTextWidthForOptions(marker, markerFont, markerNamedFont, markerSize, currentOpts));
            }

            double markerWidth = GetListMarkerWidth(style, estimatedMarkerWidth);
            double spaceAdvance = EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts);
            double markerGap = GetListMarkerGap(style, spaceAdvance);
            double listLeftIndent = style?.LeftIndent ?? 0D;
            double rawTextWidth = frameWidth - listLeftIndent - markerWidth - markerGap;
            double availableWidth = System.Math.Max(
                rawTextWidth,
                EstimateSimpleTextWidthForOptions("WW", baseFont, size, currentOpts));
            double alignmentWidth = System.Math.Max(0D, rawTextWidth);
            double itemSpacing = GetListItemSpacing(style, leading);
            var items = new System.Collections.Generic.List<PreparedListItem>(block.RichItems.Count);
            for (int itemIndex = 0; itemIndex < block.RichItems.Count; itemIndex++) {
                PdfListItem item = block.RichItems[itemIndex];
                TableCellTextLayout textLayout = CreateListItemTextLayout(
                    item,
                    availableWidth,
                    baseFont,
                    size,
                    leading,
                    currentOpts);
                double firstLineWidth = textLayout.Lines.Count > 0
                    ? MeasureRichLineWidth(textLayout.Lines[0], currentOpts)
                    : 0D;
                double firstLineOffset = block.Align switch {
                    PdfAlign.Center => System.Math.Max(0D, (alignmentWidth - firstLineWidth) / 2D),
                    PdfAlign.Right => System.Math.Max(0D, alignmentWidth - firstLineWidth),
                    _ => 0D
                };
                items.Add(new PreparedListItem(item, markers[itemIndex], textLayout, firstLineOffset));
            }

            double spacingBefore = style?.SpacingBefore ?? 0D;
            if (topLevelSpacing) {
                spacingBefore = ResolveTopLevelSpacingBefore(spacingBefore);
            }

            return new PreparedListLayout(
                block,
                style,
                markerFont,
                markerNamedFont,
                markerSize,
                markerWidth,
                markerGap,
                listLeftIndent,
                alignmentWidth,
                size,
                leading,
                itemSpacing,
                spacingBefore,
                style?.GetSpacingAfter(itemSpacing) ?? itemSpacing,
                items);
        }

        private static double MeasurePreparedListHeight(PreparedListLayout prepared) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < prepared.Items.Count; itemIndex++) {
                TableCellTextLayout textLayout = prepared.Items[itemIndex].TextLayout;
                total += itemIndex == 0 ? prepared.SpacingBefore : 0D;
                total += MeasureRichLinesHeight(textLayout.LineHeights, textLayout.LineCount, prepared.Leading);
                total += itemIndex == prepared.Items.Count - 1 ? prepared.SpacingAfter : prepared.ItemSpacing;
            }

            return total;
        }
    }
}
