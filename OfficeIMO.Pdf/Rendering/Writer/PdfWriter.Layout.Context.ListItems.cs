namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderListItem(PreparedListLayout prepared, int itemIndex, IList<IPdfBlock> blockList, int blockIndex,
            ref int? listStructureElementIndex, ref LayoutResult.Page? listStructurePage) {
            PreparedListItem preparedItem = prepared.Items[itemIndex];
            PdfListItem item = preparedItem.Item;
            var runs = item.Runs;
            var lines = preparedItem.TextLayout.Lines;
            var lineHeights = preparedItem.TextLayout.LineHeights;
            string marker = preparedItem.Marker;
            PdfStandardFont markerFont = prepared.MarkerFont;
            PdfNamedFontFace? markerNamedFont = prepared.MarkerNamedFont;
            double markerSize = prepared.MarkerSize;
            PdfColor? color = prepared.Block.Color ?? prepared.Style?.Color;
            PdfColor? markerColor = prepared.Style?.MarkerColor ?? color;
            double markerX = currentOpts.MarginLeft + prepared.ListLeftIndent + preparedItem.FirstLineOffset;
            double markerWidth = prepared.MarkerWidth;
            PdfAlign markerAlign = prepared.Block.GetMarkerAlign(prepared.Style);
            double textX = currentOpts.MarginLeft + prepared.ListLeftIndent + prepared.MarkerWidth + prepared.MarkerGap;
            double textWidth = prepared.AlignmentWidth;
            PdfAlign textAlign = prepared.Block.Align;
            double size = prepared.Size;
            double leading = prepared.Leading;
            double spacingBefore = itemIndex == 0 ? prepared.SpacingBefore : 0D;
            double spacingAfter = itemIndex == prepared.Items.Count - 1 ? prepared.SpacingAfter : prepared.ItemSpacing;
            string? bookmarkName = item.BookmarkName;
            PdfLineSpacing? lineSpacing = prepared.Style?.LineSpacing;

            double markerOffset = markerX - currentOpts.MarginLeft;
            double textOffset = textX - currentOpts.MarginLeft;
            double rightInset = width - textOffset - textWidth;
            double wrappedWidth = width;
            int lineIndex = 0;
            bool firstSegment = true;
            var listFont = ChooseNormal(currentOpts.DefaultFont);
            void NewListFrame() {
                if (CanQueueColumnBalanceRemainder(blockList)) {
                    QueueColumnBalanceRemainder(MeasureListColumnBalanceUnits(prepared, itemIndex, lineIndex, lineHeights,
                        includeSpacingBefore: false), prepared.Block, blockList, blockIndex);
                }
                NewPage();
                if (activeColumnFlow != null && Math.Abs(wrappedWidth - width) > 0.001D) {
                    textWidth = width - textOffset - rightInset;
                    var remainingRuns = firstSegment ? runs : BuildTextRunsFromWrappedLines(lines, lineIndex, lines.Count - lineIndex);
                    (lines, lineHeights) = WrapRichRunsWithSpacing(remainingRuns, textWidth, size, listFont, leading, null,
                        DefaultParagraphTabStopWidth, currentOpts, lineSpacing);
                    lineIndex = 0;
                    wrappedWidth = width;
                }
            }
            PageStructElement? listItemElement = null;
            spacingBefore = ResolveTopLevelSpacingBefore(spacingBefore);
            if (spacingBefore > 0) {
                if (y - spacingBefore < currentOpts.MarginBottom) {
                    NewListFrame();
                    spacingBefore = 0D;
                }

                if (spacingBefore > 0) y -= spacingBefore;
            }

            while (lineIndex < lines.Count) {
                double available = y - currentOpts.MarginBottom;
                double firstLineHeight = GetRichLineHeight(lineHeights, lineIndex, leading);
                if (firstLineHeight > GetMaximumBlockContinuationHeight() + .001D)
                    throw new ArgumentException("List line height exceeds the available page content height.");
                while (available + .001D < firstLineHeight) {
                    if (!ShouldAdvanceForBlockHeight(firstLineHeight))
                        throw new ArgumentException("List line height exceeds the available page content height.");
                    NewListFrame();
                    available = y - currentOpts.MarginBottom;
                    firstLineHeight = GetRichLineHeight(lineHeights, lineIndex, leading);
                    if (firstLineHeight > GetMaximumBlockContinuationHeight() + .001D) {
                        throw new ArgumentException("List line height exceeds the available page content height.");
                    }
                }

                int take = 0;
                double heightSum = 0;
                for (int k = lineIndex; k < lines.Count; k++) {
                    double lineHeight = GetRichLineHeight(lineHeights, k, leading);
                    if (heightSum + lineHeight > available) {
                        break;
                    }

                    heightSum += lineHeight;
                    take++;
                }

                ReserveBalancedKeepNextTail(lineHeights, lineIndex, lines.Count, spacingAfter, available, blockList, blockIndex,
                    itemIndex == prepared.Items.Count - 1 && prepared.Style?.KeepWithNext == true, 1, ref take, ref heightSum);
                if (take == 0) {
                    NewListFrame();
                    continue;
                }

                var segmentLines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>(take);
                var segmentHeights = new System.Collections.Generic.List<double>(take);
                for (int k = 0; k < take; k++) {
                    segmentLines.Add(lines[lineIndex + k]);
                    segmentHeights.Add(GetRichLineHeight(lineHeights, lineIndex + k, leading));
                }

                double baselineY = FirstTextBaselineFromTop(listFont, size, y);
                int? listElementIndex = firstSegment ? EnsurePageStructureContainer("L", ref listStructureElementIndex, ref listStructurePage) : null;
                int? listItemElementIndex = firstSegment ? RegisterStructureContainer("LI", listElementIndex) : null;
                if (firstSegment) {
                    if (listItemElementIndex.HasValue && currentPage != null) {
                        listItemElement = currentPage.StructElements[listItemElementIndex.Value];
                    }

                    if (!string.IsNullOrEmpty(bookmarkName)) {
                        AddNamedDestinationName(bookmarkName!, y);
                    }

                    var markerLines = new System.Collections.Generic.List<string>(1) { marker };
                    int? labelMarkedContentId = RegisterTextStructureElement("Lbl", listItemElementIndex);
                    if (markerNamedFont.HasValue) {
                        currentPage!.UsedNamedFonts.Add(markerNamedFont.Value);
                    } else {
                        MarkSimpleFont(markerFont);
                    }

                    WriteLinesInternal(
                        GetFontResourceName(markerFont, markerNamedFont, ChooseNormal(currentOpts.DefaultFont)),
                        markerSize,
                        leading,
                        currentOpts.MarginLeft + markerOffset,
                        markerWidth,
                        AdjustRichLineBaseline(baselineY, segmentLines[0], currentOpts, size),
                        markerLines,
                        markerAlign,
                        markerColor ?? color,
                        applyBaselineTweak: true,
                        structureType: "Lbl",
                        markedContentId: labelMarkedContentId,
                        namedFont: markerNamedFont);
                }

                RecordFlowPlacement(y);
                pageDirty = true;
                int? bodyMarkedContentId = firstSegment || listItemElement == null
                    ? RegisterTextStructureElement("LBody", listItemElementIndex)
                    : RegisterTextStructureElement("LBody", listItemElement);
                WriteRichParagraph(sb, new RichParagraphBlock(runs, textAlign, color), segmentLines, segmentHeights, currentOpts, baselineY, size, leading, currentPage!.Annotations, currentOpts.MarginLeft + textOffset, textWidth, structureType: "LBody", markedContentId: bodyMarkedContentId, structurePage: currentPage);
                MarkRichFonts(runs);
                y -= heightSum;
                lineIndex += take;
                firstSegment = false;
                if (lineIndex < lines.Count) {
                    NewListFrame();
                } else {
                    y -= spacingAfter;
                }
            }
        }

    }
}
