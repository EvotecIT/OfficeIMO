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
                double firstLineOffset,
                ListMarkerFallbackLayout? markerFallback) {
                Item = item;
                Marker = marker;
                TextLayout = textLayout;
                FirstLineOffset = firstLineOffset;
                MarkerFallback = markerFallback;
            }

            internal PdfListItem Item { get; }
            internal string Marker { get; }
            internal TableCellTextLayout TextLayout { get; }
            internal double FirstLineOffset { get; }
            internal ListMarkerFallbackLayout? MarkerFallback { get; }
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
            var markerFallbacks = new ListMarkerFallbackLayout?[markers.Length];
            // Estimate from the actual markers, including default/custom mixtures.
            // The selected font need not contain unused numbering or bullet glyphs.
            double estimatedMarkerWidth = 0D;
            for (int itemIndex = 0; itemIndex < markers.Length; itemIndex++) {
                string marker = block.GetMarker(itemIndex);
                markers[itemIndex] = marker;
                markerFallbacks[itemIndex] = PrepareListMarkerFallback(marker, markerFont, markerSize, leading, style);
                estimatedMarkerWidth = System.Math.Max(
                    estimatedMarkerWidth,
                    markerFallbacks[itemIndex]?.Width ??
                        EstimateSimpleTextWidthForOptions(marker, markerFont, markerNamedFont, markerSize, currentOpts, style?.MarkerFeatureSettings));
            }

            double markerWidth = GetListMarkerWidth(style, estimatedMarkerWidth);
            double spaceAdvance = style?.MarkerGap.HasValue == true ? 0D :
                EstimateSimpleTextWidthForOptions(" ", markerFont, markerNamedFont, markerSize, currentOpts, style?.MarkerFeatureSettings);
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
                    currentOpts,
                    style?.LineSpacing,
                    cancellationToken);
                double firstLineWidth = textLayout.Lines.Count > 0
                    ? MeasureRichLineWidth(textLayout.Lines[0], currentOpts)
                    : 0D;
                double firstLineOffset = block.Align switch {
                    PdfAlign.Center => System.Math.Max(0D, (alignmentWidth - firstLineWidth) / 2D),
                    PdfAlign.Right => System.Math.Max(0D, alignmentWidth - firstLineWidth),
                    _ => 0D
                };
                items.Add(new PreparedListItem(item, markers[itemIndex], textLayout, firstLineOffset, markerFallbacks[itemIndex]));
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

        private (double X, double Width, PdfAlign Align) GetListMarkerPlacement(
            PreparedListLayout prepared, PreparedListItem item) {
            double x = prepared.ListLeftIndent + item.FirstLineOffset;
            PdfAlign align = prepared.Block.GetMarkerAlign(prepared.Style);
            if (prepared.Style?.MarkerAlignsAtIndent != true)
                return (x, prepared.MarkerWidth, align);

            double width = item.MarkerFallback?.Width ?? EstimateSimpleTextWidthForOptions(item.Marker, prepared.MarkerFont,
                prepared.MarkerNamedFont, prepared.MarkerSize, currentOpts, prepared.Style?.MarkerFeatureSettings);
            x -= align switch {
                PdfAlign.Right => width,
                PdfAlign.Center => width / 2D,
                _ => 0D
            };
            return (x, width, PdfAlign.Left);
        }

        private static double MeasurePreparedListHeight(PreparedListLayout prepared, bool frameFit = false) {
            double total = 0D;
            for (int itemIndex = 0; itemIndex < prepared.Items.Count; itemIndex++) {
                TableCellTextLayout textLayout = prepared.Items[itemIndex].TextLayout;
                total += itemIndex == 0 ? prepared.SpacingBefore : 0D;
                total += frameFit && itemIndex == prepared.Items.Count - 1
                    ? MeasureRichLinesFrameFitHeight(textLayout.Lines, textLayout.LineHeights, prepared.Leading)
                    : MeasureRichLinesHeight(textLayout.LineHeights, textLayout.LineCount, prepared.Leading);
                total += itemIndex == prepared.Items.Count - 1 ? prepared.SpacingAfter : prepared.ItemSpacing;
            }

            return total;
        }
    }
}
