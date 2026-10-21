namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints the visible text and objects of one admitted cell fragment; links share its bounded frame.</summary>
        private int RenderFlowTableCellContent(TableBlock table, PdfTableStyle style, TableCellLayout cell,
            TableCellTextLayout lines, int rowIndex, int sourceStartLine, int requestedLineCount,
            double cellX, double cellTop, double cellWidth, double cellHeight, double rowSize,
            double rowLeading, double runFontSizeScale, bool rowUsesBold, bool renderAsHeader,
            bool wholeRowSegment, bool suppressCellObjects, PdfColor? textColor,
            int? rowStructureElementIndex, PageStructElement? rowStructureElement = null,
            int? fragmentRowSpan = null, bool includeNamedDestination = true, double? verticalOffsetOverride = null) {
            int c = cell.Column;
            double textClipBleed = style.ClipTextToCellBounds ? 0D : TableCellClipBleed;
            double cellPadLeft = GetTableCellPaddingLeft(style, rowIndex, c);
            double cellPadRight = GetTableCellPaddingRight(style, rowIndex, c);
            FlowTableCellTextFrame textFrame = PrepareFlowTableCellTextFrame(style, cell, lines, rowIndex,
                sourceStartLine, requestedLineCount, cellX, cellTop, cellWidth, cellHeight, rowSize,
                rowLeading, rowUsesBold, verticalOffsetOverride);
            TableCellContentFrame contentFrame = textFrame.Content;
            double innerW = textFrame.InnerWidth;
            double cellBottom = cellTop - cellHeight;
            PdfColumnAlign align = GetTableCellAlignment(style, rowIndex, c, cell.Text);
            var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
            // A neighboring horizontal cell can continue the row after the
            // turned cell's content has already been painted on its first fragment.
            int visibleLineCount = textFrame.VisibleLineCount;
            double verticalOffset = textFrame.VerticalOffset;
            double visibleTextHeight = textFrame.VisibleTextHeight;

            // Deferred row spans consume only the lines actually admitted and resume the rest.
            if (!verticalOffsetOverride.HasValue)
                ReportFlowTableCellDiagnostics(style, cell, lines, rowIndex, sourceStartLine, requestedLineCount,
                    visibleLineCount, cellX, cellTop, cellWidth, cellHeight, rowLeading, innerW,
                    verticalOffset, includeObjects: !suppressCellObjects);

            double firstBaseline = textFrame.FirstBaseline;

            pageDirty = true;
            MarkRichFonts(cell.Runs, forceBold: rowUsesBold);
            string? linkUri = cell.LinkUri;
            string? linkDestinationName = cell.LinkDestinationName;
            string? linkContents = cell.LinkContents;
            if (table.Links.TryGetValue((rowIndex, c), out var uri)) {
                linkUri = uri;
                linkDestinationName = null;
                linkContents = cell.Text;
            }

            if (sourceStartLine == 0 && includeNamedDestination) {
                AddTableCellNamedDestinationName(cell.NamedDestinationName, cellTop);
            }

            int? cellLinkStructElementIndex = null;
            if (visibleLineCount > 0) {
                var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                var visibleAlignments = SliceTableCellLineAlignments(lines, sourceStartLine, visibleLineCount);
                var visibleXOffsets = SliceTableCellLineXOffsets(lines, sourceStartLine, visibleLineCount);
                var visibleWidths = SliceTableCellLineWidths(lines, sourceStartLine, visibleLineCount, innerW);
                TableCellContentFrame inkClip = GetTableCellInkClipFrame(style, rowIndex, c, cellX, cellTop, cellWidth, cellHeight);
                double textClipX = inkClip.Left - textClipBleed;
                double textClipWidth = inkClip.Width + (textClipBleed * 2D);
                if (cell.Viewport == null && !style.ClipTextToCellBounds) ExpandTableCellTextClip(cellX + cellPadLeft, visibleXOffsets, visibleWidths, ref textClipX, ref textClipWidth);
                var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                string structureType = renderAsHeader ? "TH" : "TD";
                int tableColumnSpan = cell.ColumnSpan > 1 ? cell.ColumnSpan : 1;
                int tableRowSpan = fragmentRowSpan ?? (wholeRowSegment && cell.RowSpan > 1 ? cell.RowSpan : 1);
                bool cellHasLinkTarget = HasCellLinkTarget(linkUri, linkDestinationName);
                int? markedContentId;
                string markedStructureType = structureType;
                if (cellHasLinkTarget && emitGeneratedStructure && currentPage != null) {
                    PageStructElement? cellElement = rowStructureElement == null
                        ? null
                        : RegisterStructureContainer(structureType, rowStructureElement, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan, logicalOrder: c);
                    int? cellElementIndex = cellElement == null
                        ? RegisterStructureContainer(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan, logicalOrder: c)
                        : null;
                    markedStructureType = "Link";
                    markedContentId = cellElement == null
                        ? RegisterTextStructureElement(markedStructureType, cellElementIndex)
                        : RegisterTextStructureElement(markedStructureType, cellElement);
                    cellLinkStructElementIndex = FindStructElementIndex(currentPage, markedContentId, markedStructureType);
                } else {
                    markedContentId = rowStructureElement == null
                        ? RegisterTextStructureElement(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan, logicalOrder: c)
                        : RegisterTextStructureElement(structureType, rowStructureElement, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan, logicalOrder: c);
                }

                if (cell.TextRotation == 0)
                    ReportRichPaintDiagnostics("TableCell", "PdfTableCell[" + rowIndex + "," + c + "]", paragraph,
                        visibleLines, visibleHeights, firstBaseline, rowSize, rowLeading, contentFrame.Left + cellPadLeft, innerW,
                        new TableCellContentFrame(textClipX, inkClip.Top + textClipBleed, textClipWidth, inkClip.Height + 2D * textClipBleed),
                        alignments: visibleAlignments, offsets: visibleXOffsets, widths: visibleWidths, baselineFont: cellFont);
                if (cell.Viewport != null || style.PreservePartialCellLines)
                    OmitInvisibleTableCellViewportLines(visibleLines, visibleHeights, visibleAlignments, visibleXOffsets, visibleWidths,
                        paragraph.Align, firstBaseline, contentFrame.Left + cellPadLeft, innerW,
                        inkClip.Left, inkClip.Top - inkClip.Height, inkClip.Width, inkClip.Height, rowLeading, rowSize, currentOpts, cellFont);
                if (cell.TextRotation != 0)
                    RenderOrientedTableCellContent(cell, style, rowIndex, c, contentFrame, cellX, cellTop, cellWidth, cellHeight, cellFont, rowSize, rowLeading, runFontSizeScale, paragraph, markedStructureType, markedContentId, includeCellObjects: !suppressCellObjects);
                else
                    WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, textClipX, inkClip.Top - inkClip.Height - textClipBleed, textClipWidth, inkClip.Height + (textClipBleed * 2D), contentFrame.Left + cellPadLeft, innerW, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, lineAlignments: visibleAlignments, lineXOffsets: visibleXOffsets, lineWidths: visibleWidths, baselineFont: cellFont);
            }
            if (cell.TextRotation == 0 && !suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                    RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], cellX + cellPadLeft, innerW, AdjustRichLineBaseline(firstBaseline, lines.Lines[sourceStartLine], currentOpts, rowSize, cellFont));
                } else {
                    double formFieldTop = textFrame.ContentTop - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                    TableCellContentFrame? clip = cell.Viewport == null ? null : new TableCellContentFrame(cellX, cellTop, cellWidth, cellHeight);
                    RenderTableCellObjects(currentPage!, cell, align, contentFrame.Left + cellPadLeft, innerW, formFieldTop,
                        clip.HasValue ? image => WriteTableCellViewportImage(image, clip) : null, clip);
                }
            }

            if (HasCellLinkTarget(linkUri, linkDestinationName) && (cell.TextRotation == 0 || sourceStartLine == 0)) {
                if (visibleLineCount == 0 && emitGeneratedStructure && currentPage != null) {
                    string structureType = renderAsHeader ? "TH" : "TD";
                    int rowSpan = fragmentRowSpan ?? (wholeRowSegment ? cell.RowSpan : 1);
                    if (rowStructureElement == null) {
                        int? cellIndex = RegisterStructureContainer(structureType, rowStructureElementIndex,
                            renderAsHeader ? "Column" : string.Empty, cell.ColumnSpan, rowSpan, logicalOrder: c);
                        cellLinkStructElementIndex = RegisterStructureContainer("Link", cellIndex);
                    } else {
                        PageStructElement? cellElement = RegisterStructureContainer(structureType, rowStructureElement,
                            renderAsHeader ? "Column" : string.Empty, cell.ColumnSpan, rowSpan, logicalOrder: c);
                        PageStructElement? linkElement = RegisterStructureContainer("Link", cellElement);
                        if (linkElement != null) cellLinkStructElementIndex = currentPage.StructElements.IndexOf(linkElement);
                    }
                }
                double x1 = cellX + cellPadLeft - textClipBleed;
                double x2 = cellX + cellWidth - cellPadRight + textClipBleed;
                double linkCellHeight = cellHeight;
                double y1 = cellTop - linkCellHeight - textClipBleed;
                double y2 = cellTop + textClipBleed;
                currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text, StructElementIndex = cellLinkStructElementIndex });
            }
            return visibleLineCount;
        }
    }
}
