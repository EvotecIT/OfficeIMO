using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class ColumnTableCursor {
            public int Index, Line, Subline;
            public double Y, Remaining, Consumed;
        }

        private bool RenderColumnTable(ColTable table, List<ColItem> items, ColumnTableCursor state, double xCol, double wCol, double fullColumnHeight, double columnPageStartY) {
        var tbColumn = table.Block;
        var tableStyle = table.Style;
        double padLeft = GetTableCellPaddingLeft(tableStyle);
        double padRight = GetTableCellPaddingRight(tableStyle);
        double padTop = GetTableCellPaddingTop(tableStyle);
        double padBottom = GetTableCellPaddingBottom(tableStyle);
        double columnGap = GetTableCellSpacing(tableStyle);
        double columnTableRowGap = columnGap;
        double xTable = ResolveTableX(tbColumn.Align, tableStyle, xCol, wCol, table.Width);

        double maxContentHeight = fullColumnHeight;
        double tableSpacingBefore = state.Line == 0 && state.Consumed > 0.001 ? tableStyle.SpacingBefore : 0D;
        if (state.Line == 0 && tableStyle.KeepTogether) {
            double keepHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
            if (keepHeight > maxContentHeight + 0.001) {
                throw new ArgumentException("Table height exceeds the available page content height.");
            }

            if (keepHeight > state.Remaining + 0.001) {
                if (state.Consumed > 0) return false;
                state.Remaining = 0;
                return false;
            }
        }

        if (state.Line == 0 && tableStyle.KeepWithNext && state.Index + 1 < items.Count) {
            double tableHeight = tableSpacingBefore + table.CaptionHeight + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, columnTableRowGap) + tableStyle.SpacingAfter;
            double nextHeight = MeasureColKeepWithNextChainHeight(items, state.Index + 1);
            double keepHeight = tableHeight + nextHeight;
            if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && keepHeight > state.Remaining + 0.001) {
                if (state.Consumed > 0) return false;
                state.Remaining = 0;
                return false;
            }
        }

        if (state.Line == 0 && state.Consumed > 0.001) {
            int minimumFirstPageBodyRows = Math.Min(
                tableStyle.MinimumBodyRowsOnFirstPage,
                Math.Max(0, table.FooterStartRowIndex - table.HeaderRowCount));
            if (minimumFirstPageBodyRows > 0) {
                int firstPageRowCount = table.HeaderRowCount + minimumFirstPageBodyRows;
                double firstPageGroupHeight =
                    tableSpacingBefore +
                    table.CaptionHeight +
                    GetTableRowsHeight(table.RowHeights, 0, firstPageRowCount, columnTableRowGap);
                if (firstPageGroupHeight <= maxContentHeight + 0.001 &&
                    firstPageGroupHeight > state.Remaining + 0.001) {
                    return false;
                }
            }
        }

        if (state.Line == 0 && tableSpacingBefore > 0) {
            if (tableSpacingBefore > state.Remaining && state.Consumed > 0) return false;
            if (tableSpacingBefore > state.Remaining && state.Consumed == 0) { state.Remaining = 0; return false; }
            state.Y -= tableSpacingBefore;
            state.Remaining -= tableSpacingBefore;
            state.Consumed += tableSpacingBefore;
        }

        int? tableStructureElementIndex = null;
        LayoutResult.Page? tableStructurePage = null;
        int? EnsureTableStructureElement() {
            if (!emitGeneratedStructure || currentPage == null) {
                return null;
            }

            if (!ReferenceEquals(tableStructurePage, currentPage)) {
                tableStructurePage = currentPage;
                tableStructureElementIndex = RegisterStructureContainer("Table", alternativeText: tableStyle.AlternativeText);
            }

            return tableStructureElementIndex;
        }

        if (state.Line == 0 && table.CaptionRuns != null && table.CaptionLines != null && table.CaptionLineHeights != null) {
            double firstRowHeight = table.RowHeights.Length > 0 ? table.RowHeights[0] : 0;
            double neededWithFirstRow = table.CaptionHeight + firstRowHeight;
            if (neededWithFirstRow > maxContentHeight + 0.001) {
                throw new ArgumentException("Table caption and first row exceed the available page content height.");
            }
            if (neededWithFirstRow > state.Remaining && state.Consumed > 0) return false;
            if (neededWithFirstRow > state.Remaining && state.Consumed == 0) { state.Remaining = 0; return false; }

            double captionSize = tableStyle.CaptionFontSize ?? table.Size;
            var captionFont = ChooseNormal(currentOpts.DefaultFont);
            pageDirty = true;
            int? captionMarkedContentId = RegisterTextStructureElement("Caption", EnsureTableStructureElement());
            MarkRichFonts(table.CaptionRuns);
            WriteRichParagraph(sb, new RichParagraphBlock(table.CaptionRuns, tableStyle.CaptionAlign, tableStyle.CaptionColor), table.CaptionLines, table.CaptionLineHeights, currentOpts, FirstTextBaselineFromTop(captionFont, captionSize, state.Y), captionSize, table.CaptionLeading, currentPage!.Annotations, xTable, table.Width, structureType: "Caption", markedContentId: captionMarkedContentId, structurePage: currentPage);
            state.Y -= table.CaptionHeight;
            state.Remaining -= table.CaptionHeight;
            state.Consumed += table.CaptionHeight;
        }

        double repeatHeaderHeight = 0;
        for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
            repeatHeaderHeight += table.RowHeights[headerIndex] + GetTableRowGapAfter(headerIndex, tbColumn.Rows.Count, columnTableRowGap);
        }

        bool HasRepeatableHeader() =>
            table.RepeatHeaderRowCount > 0 &&
            tbColumn.Rows.Count > table.HeaderRowCount;

        bool AtContinuationPageTop() =>
            Math.Abs(state.Y - columnPageStartY) <= 0.001;

        double MeasureColumnTableRowSegmentHeight(int rowIndex, int startLine, int lineCount, bool suppressCellObjects) {
            double rowLeading = table.RowLeadings[rowIndex];
            double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
            double segmentHeight = rowLeading + rowPadTop + rowPadBottom;
            var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double cellWidth = GetTableCellWidth(table.ColumnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double cellPadLeft = GetTableCellPaddingLeft(tableStyle, rowIndex, cell.Column);
                double cellPadRight = GetTableCellPaddingRight(tableStyle, rowIndex, cell.Column);
                double innerW = cellWidth - cellPadLeft - cellPadRight;
                TableCellTextLayout lines = table.RowLines[rowIndex][cell.Column];
                int sourceStartLine = startLine;
                int visibleLineCount = Math.Max(0, Math.Min(lineCount, lines.LineCount - sourceStartLine));
                bool includeObjects = !suppressCellObjects && sourceStartLine == 0;
                double cellContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW, includeObjects) +
                    GetTableCellPaddingTop(tableStyle, rowIndex, cell.Column) +
                    GetTableCellPaddingBottom(tableStyle, rowIndex, cell.Column);
                segmentHeight = Math.Max(segmentHeight, cellContentHeight);
            }

            return segmentHeight;
        }

        int GetColumnTableRowSegmentLineCountThatFits(int rowIndex, int startLine, double available) {
            int remainingLines = table.RowLineCounts[rowIndex] - startLine;
            int best = 0;
            for (int candidate = 1; candidate <= remainingLines; candidate++) {
                double candidateHeight = MeasureColumnTableRowSegmentHeight(rowIndex, startLine, candidate, suppressCellObjects: false);
                if (candidateHeight > available + 0.001) {
                    break;
                }

                best = candidate;
            }

            return Math.Max(1, best);
        }

        bool CanSplitColumnTableRowIntoRemainingSpace(int rowIndex) =>
            rowIndex >= table.HeaderRowCount &&
            GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex) &&
            table.RowLineCounts[rowIndex] > 1 &&
            MeasureColumnTableRowSegmentHeight(rowIndex, 0, Math.Min(2, table.RowLineCounts[rowIndex]), suppressCellObjects: false) <= state.Remaining + 0.001;

        bool ShouldBreakBeforeFinalColumnTableBodyRows(int rowIndex) {
            int minimumBodyRows = Math.Min(tableStyle.MinimumBodyRowsOnLastPage, Math.Max(0, table.FooterStartRowIndex - table.HeaderRowCount));
            if (minimumBodyRows <= 0 || table.FooterStartRowIndex - rowIndex != minimumBodyRows) {
                return false;
            }

            double currentRowHeight = table.RowHeights[rowIndex] + GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
            double finalGroupHeight = GetTableRowsHeight(table.RowHeights, rowIndex, table.RowHeights.Length, columnTableRowGap);
            return ShouldBreakBeforeFinalTableBodyRows(
                rowIndex,
                table.HeaderRowCount,
                table.FooterStartRowIndex,
                minimumBodyRows,
                currentRowHeight,
                finalGroupHeight,
                state.Remaining,
                HasRepeatableHeader() ? repeatHeaderHeight : 0D,
                maxContentHeight,
                state.Consumed > 0.001);
        }

        void DrawColumnTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false) {
            bool renderAsFooter = rowIndex >= table.FooterStartRowIndex;
            bool rowUsesBold = table.RowBold[rowIndex];
            double rowSize = table.RowSizes[rowIndex];
            double rowLeading = table.RowLeadings[rowIndex];
            bool wholeRowSegment = startLine == 0 && lineCount == table.RowLineCounts[rowIndex];
            double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowHeight = wholeRowSegment ? table.RowHeights[rowIndex] : MeasureColumnTableRowSegmentHeight(rowIndex, startLine, lineCount, suppressCellObjects);
            if (rowUsesBold) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
            double rowBottom = state.Y - rowHeight;
            int bodyRowIndex = rowIndex - table.HeaderRowCount;
            bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
            bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
            if (tableStyle.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, tableStyle.HeaderFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
            else if (tableStyle.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, tableStyle.FooterFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
            else if (!renderAsHeader && !renderAsFooter && tableStyle.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, tableStyle.RowStripeFill.Value, xTable, table.ColumnWidths, columnGap, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }

            if (!renderAsHeader && !renderAsFooter && tableStyle.BodyColumnFills != null) {
                bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
                double fillX = xTable;
                for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                    PdfColor? fill = fillColumn < tableStyle.BodyColumnFills.Count ? tableStyle.BodyColumnFills[fillColumn] : null;
                    if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                        pageDirty = true;
                        DrawRowFill(sb, fill.Value, fillX, rowBottom, table.ColumnWidths[fillColumn], rowHeight, emitGeneratedStructure);
                    }
                    fillX += table.ColumnWidths[fillColumn] + columnGap;
                }
            }

            if (tableStyle.CellFills != null && tableStyle.CellFills.Count > 0) {
                double fillX = xTable;
                for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                    if (tableStyle.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                        TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                        (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                        int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                        double fillHeight = rowHeight;
                        double fillBottom = rowBottom;
                        if (wholeRowSegment) {
                            if (fillCell.RowSpan > 1) {
                                fillHeight = GetTableCellHeight(table.RowHeights, rowIndex, fillCell.RowSpan, columnTableRowGap);
                                fillBottom = state.Y - fillHeight;
                            }
                        }

                        pageDirty = true;
                        DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, columnGap), fillHeight, emitGeneratedStructure);
                    }
                    fillX += table.ColumnWidths[fillColumn] + columnGap;
                }
            }
            if (DrawTableCellDataBars(sb, tableStyle, cells, rowIndex, table.Columns, xTable, state.Y, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                pageDirty = true;
            }
            if (DrawTableCellIcons(sb, tableStyle, cells, rowIndex, table.Columns, xTable, state.Y, rowBottom, rowHeight, table.ColumnWidths, columnGap, table.RowHeights, columnTableRowGap, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                pageDirty = true;
            }

            var textColor = renderAsHeader ? tableStyle.HeaderTextColor : renderAsFooter ? tableStyle.FooterTextColor : tableStyle.TextColor;
            double xi = xTable;
            int? rowStructureElementIndex = RegisterStructureContainer("TR", EnsureTableStructureElement());
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                int c = cell.Column;
                xi = xTable;
                for (int xColumn = 0; xColumn < c; xColumn++) {
                    xi += table.ColumnWidths[xColumn] + columnGap;
                }

                double cellWidth = GetTableCellWidth(table.ColumnWidths, c, cell.ColumnSpan, columnGap);
                double cellPadLeft = GetTableCellPaddingLeft(tableStyle, rowIndex, c);
                double cellPadRight = GetTableCellPaddingRight(tableStyle, rowIndex, c);
                double cellPadTop = GetTableCellPaddingTop(tableStyle, rowIndex, c);
                double cellPadBottom = GetTableCellPaddingBottom(tableStyle, rowIndex, c);
                double innerW = cellWidth - cellPadLeft - cellPadRight;
                double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap) : rowHeight;
                double cellBottom = state.Y - cellHeight;
                PdfColumnAlign align = GetTableCellAlignment(tableStyle, rowIndex, c, cell.Text);
                PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(tableStyle, rowIndex, c);
                var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                TableCellTextLayout lines = table.RowLines[rowIndex][c];
                int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                double availableTextHeight = Math.Max(0, cellHeight - cellPadTop - cellPadBottom);
                int visibleLineCount = LimitTableCellLineCountToHeight(lines, sourceStartLine, requestedLineCount, rowLeading, availableTextHeight);
                double verticalOffset = 0;
                double visibleTextHeight = 0D;
                if (visibleLineCount > 0) {
                    visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleLineCount, rowLeading);
                    double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW);
                    double unusedTextHeight = Math.Max(0, availableTextHeight - visibleContentHeight);
                    if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                    else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                }

                double firstBaseline = state.Y - cellPadTop - verticalOffset - GetAscenderForOptions(cellFont, rowSize, currentOpts) + tableStyle.RowBaselineOffset;

                pageDirty = true;
                if (cell.Runs.Any(run => run.Bold || rowUsesBold)) { currentPage!.UsedBold = true; usedBold = true; }
                if (cell.Runs.Any(run => run.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                MarkRichFonts(cell.Runs);
                string? linkUri = cell.LinkUri;
                string? linkDestinationName = cell.LinkDestinationName;
                string? linkContents = cell.LinkContents;
                if (tbColumn.Links.TryGetValue((rowIndex, c), out var uri)) {
                    linkUri = uri;
                    linkDestinationName = null;
                    linkContents = cell.Text;
                }

                if (sourceStartLine == 0) {
                    AddTableCellNamedDestinationName(cell.NamedDestinationName, state.Y);
                }

                int? cellLinkStructElementIndex = null;
                if (visibleLineCount > 0) {
                    var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                    visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                    var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                    var visibleAlignments = SliceTableCellLineAlignments(lines, sourceStartLine, visibleLineCount);
                    var visibleXOffsets = SliceTableCellLineXOffsets(lines, sourceStartLine, visibleLineCount);
                    var visibleWidths = SliceTableCellLineWidths(lines, sourceStartLine, visibleLineCount, innerW);
                    double textClipX = xi - TableCellClipBleed;
                    double textClipWidth = cellWidth + (TableCellClipBleed * 2D);
                    ExpandTableCellTextClip(xi + cellPadLeft, innerW, cell.NoWrap, visibleXOffsets, visibleWidths, ref textClipX, ref textClipWidth);
                    var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                    string structureType = renderAsHeader ? "TH" : "TD";
                    int tableColumnSpan = cell.ColumnSpan > 1 ? cell.ColumnSpan : 1;
                    int tableRowSpan = wholeRowSegment && cell.RowSpan > 1 ? cell.RowSpan : 1;
                    bool cellHasLinkTarget = HasCellLinkTarget(linkUri, linkDestinationName);
                    int? markedContentId;
                    string markedStructureType = structureType;
                    if (cellHasLinkTarget && emitGeneratedStructure && currentPage != null) {
                        int? cellElementIndex = RegisterStructureContainer(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                        markedStructureType = "Link";
                        markedContentId = RegisterTextStructureElement(markedStructureType, cellElementIndex);
                        cellLinkStructElementIndex = FindStructElementIndex(currentPage, markedContentId, markedStructureType);
                    } else {
                        markedContentId = RegisterTextStructureElement(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                    }

                    WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, textClipX, cellBottom - TableCellClipBleed, textClipWidth, cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, lineAlignments: visibleAlignments, lineXOffsets: visibleXOffsets, lineWidths: visibleWidths);
                }
                if (!suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                    if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                        RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], xi + cellPadLeft, innerW, firstBaseline);
                    } else {
                        double formFieldTop = state.Y - cellPadTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                        RenderTableCellObjects(currentPage!, cell, align, xi + cellPadLeft, innerW, formFieldTop);
                    }
                }

                if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                    double linkCellHeight = sourceStartLine == 0 && cell.RowSpan > 1
                        ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap)
                        : cellHeight;
                    currentPage!.Annotations.Add(new LinkAnnotation { X1 = xi + cellPadLeft - TableCellClipBleed, Y1 = state.Y - linkCellHeight - TableCellClipBleed, X2 = xi + cellWidth - cellPadRight + TableCellClipBleed, Y2 = state.Y + TableCellClipBleed, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text, StructElementIndex = cellLinkStructElementIndex });
                }
            }

            if (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) {
                pageDirty = true;
                bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns);
                bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns);
                bool segmentBorderRows = HasSkippedColumns(topBorderSkips, table.Columns) || HasSkippedColumns(bottomBorderSkips, table.Columns);
                if (segmentBorderRows) {
                    DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom + rowHeight, topBorderSkips, emitGeneratedStructure);
                    DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom, bottomBorderSkips, emitGeneratedStructure);
                    DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                    DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable + table.Width, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                } else {
                    DrawRowRect(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom, table.Width, rowHeight, emitGeneratedStructure);
                }

                double xi2 = xTable;
                for (int c = 0; c < table.Columns - 1; c++) {
                    xi2 += table.ColumnWidths[c];
                    if (IsTableBoundaryInsideSpannedCell(tbColumn, rowIndex, c, table.Columns)) {
                        xi2 += columnGap;
                        continue;
                    }

                    DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xi2, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                    xi2 += columnGap;
                }
            }

            if (renderAsFooter && rowIndex == table.FooterStartRowIndex) {
                PdfColor? footerSeparatorColor = tableStyle.FooterSeparatorColor ?? tableStyle.RowSeparatorColor;
                double footerSeparatorWidth = tableStyle.FooterSeparatorWidth > 0 ? tableStyle.FooterSeparatorWidth : tableStyle.RowSeparatorWidth;
                if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                    pageDirty = true;
                    DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xTable, table.ColumnWidths, columnGap, state.Y, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns), emitGeneratedStructure);
                }
            }

            PdfColor? separatorColor = renderAsHeader && tableStyle.HeaderSeparatorColor is not null ? tableStyle.HeaderSeparatorColor : tableStyle.RowSeparatorColor;
            double separatorWidth = renderAsHeader && tableStyle.HeaderSeparatorWidth > 0 ? tableStyle.HeaderSeparatorWidth : tableStyle.RowSeparatorWidth;
            if (separatorColor is not null && separatorWidth > 0) {
                pageDirty = true;
                DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xTable, table.ColumnWidths, columnGap, rowBottom, GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns), emitGeneratedStructure);
            }

            if (tableStyle.CellBorders != null && tableStyle.CellBorders.Count > 0) {
                double borderX = xTable;
                for (int borderColumn = 0; borderColumn < table.Columns; borderColumn++) {
                    if (tableStyle.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                        TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                        (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                        HasRenderableCellBorder(cellBorder)) {
                        int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                        double borderHeight = rowHeight;
                        double borderBottom = rowBottom;
                        if (wholeRowSegment) {
                            if (borderCell.RowSpan > 1) {
                                borderHeight = GetTableCellHeight(table.RowHeights, rowIndex, borderCell.RowSpan, columnTableRowGap);
                                borderBottom = state.Y - borderHeight;
                            }
                        }

                        pageDirty = true;
                        DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight, emitGeneratedStructure);
                    }
                    borderX += table.ColumnWidths[borderColumn] + columnGap;
                }
            }

            double rowAdvance = rowHeight + (wholeRowSegment ? GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) : 0D);
            state.Y -= rowAdvance;
            state.Remaining -= rowAdvance;
            state.Consumed += rowAdvance;
        }

        void DrawColumnTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
            DrawColumnTableRowSegment(rowIndex, renderAsHeader, 0, table.RowLineCounts[rowIndex], suppressCellObjects);

        int rowIndex = state.Line;
        int rowStartLine = state.Subline;
        while (rowIndex < tbColumn.Rows.Count) {
            double rowHeight = table.RowHeights[rowIndex];
            if (rowHeight > maxContentHeight + 0.001) {
                if (!GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex)) {
                    throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                }

                int totalLines = table.RowLineCounts[rowIndex];
                double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
                double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
                bool repeatHeaderBeforeSegment = rowIndex >= table.HeaderRowCount &&
                    HasRepeatableHeader() &&
                    AtContinuationPageTop() &&
                    repeatHeaderHeight + table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom <= state.Remaining + 0.001;
                double neededForFirstSegment = table.RowLeadings[rowIndex] + rowPadTop + rowPadBottom + (repeatHeaderBeforeSegment ? repeatHeaderHeight : 0);
                if (neededForFirstSegment > state.Remaining && state.Consumed > 0) break;
                if (neededForFirstSegment > state.Remaining && state.Consumed == 0) { state.Remaining = 0; break; }

                if (repeatHeaderBeforeSegment) {
                    for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                        DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                    }
                }

                int take = Math.Min(totalLines - rowStartLine, GetColumnTableRowSegmentLineCountThatFits(rowIndex, rowStartLine, state.Remaining));
                DrawColumnTableRowSegment(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount && rowStartLine == 0, rowStartLine, take);
                rowStartLine += take;

                if (rowStartLine < totalLines) {
                    state.Line = rowIndex;
                    state.Subline = rowStartLine;
                    break;
                }

                double gapAfterSplitRow = GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
                if (gapAfterSplitRow > 0) {
                    state.Y -= gapAfterSplitRow;
                    state.Remaining -= gapAfterSplitRow;
                    state.Consumed += gapAfterSplitRow;
                }

                rowIndex++;
                state.Line = rowIndex;
                state.Subline = 0;
                rowStartLine = 0;
                continue;
            }
            bool repeatHeaderBeforeRow = rowIndex >= table.HeaderRowCount &&
                HasRepeatableHeader() &&
                AtContinuationPageTop() &&
                repeatHeaderHeight + rowHeight <= state.Remaining + 0.001;
            double neededForNextRow = rowHeight + GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) + (repeatHeaderBeforeRow ? repeatHeaderHeight : 0);
            if (rowHeight > state.Remaining + 0.001 && state.Consumed > 0 && CanSplitColumnTableRowIntoRemainingSpace(rowIndex)) {
                int take = Math.Min(table.RowLineCounts[rowIndex], GetColumnTableRowSegmentLineCountThatFits(rowIndex, 0, state.Remaining));
                DrawColumnTableRowSegment(rowIndex, renderAsHeader: false, 0, take);
                state.Line = rowIndex;
                state.Subline = take;
                break;
            }

            if (ShouldBreakBeforeFinalColumnTableBodyRows(rowIndex)) break;
            if (neededForNextRow > state.Remaining && state.Consumed > 0) break;
            if (neededForNextRow > state.Remaining && state.Consumed == 0) { state.Remaining = 0; break; }

            if (repeatHeaderBeforeRow) {
                for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                    DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                }
            }

            DrawColumnTableRow(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount);
            rowIndex++;
            state.Line = rowIndex;
            state.Subline = 0;
            rowStartLine = 0;
        }

        if (rowIndex >= tbColumn.Rows.Count) {
            if (tableStyle.SpacingAfter > 0 && tableStyle.SpacingAfter <= state.Remaining) {
                state.Y -= tableStyle.SpacingAfter;
                state.Remaining -= tableStyle.SpacingAfter;
                state.Consumed += tableStyle.SpacingAfter;
            }
            state.Index++;
            state.Line = 0;
            state.Subline = 0;
        } else {
            return false;
        }
            return true;
        }
    }
}
