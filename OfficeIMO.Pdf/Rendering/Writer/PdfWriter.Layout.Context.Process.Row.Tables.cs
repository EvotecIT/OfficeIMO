using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class ColumnTableCursor {
            public int Index, Line, Subline;
            public double Y, Remaining, Consumed;
        }

        private bool RenderColumnTable(ColTable table, List<ColItem> items, ColumnTableCursor state, double xCol, double wCol, double fullColumnHeight, double columnPageStartY, double closingPadding = 0D) {
        var tbColumn = table.Block;
        if (closingPadding > 0D) closingPadding += table.Style.SpacingAfter;
        var tableStyle = table.Style;
        StringBuilder? pairedBorders = null;
        double textClipBleed = tableStyle.ClipTextToCellBounds ? 0D : TableCellClipBleed;
        bool tableStartedInThisColumn = state.Line == 0 && state.Subline == 0;
        double flowYBeforeTable = state.Y;
        double flowRemainingBeforeTable = state.Remaining;
        double flowConsumedBeforeTable = state.Consumed;
        double padLeft = GetTableCellPaddingLeft(tableStyle);
        double padRight = GetTableCellPaddingRight(tableStyle);
        double padTop = GetTableCellPaddingTop(tableStyle);
        double padBottom = GetTableCellPaddingBottom(tableStyle);
        double columnGap = GetTableCellSpacing(tableStyle);
        double columnTableRowGap = columnGap;
        double xTable = ResolveTableX(tbColumn.Align, tableStyle, xCol, wCol, table.Width);
        double tableCornerRadius = ResolveTableCornerRadius(
            tableStyle.CornerRadius,
            table.Width,
            table.RowHeights[0],
            table.RowHeights[table.RowHeights.Length - 1],
            table.ColumnWidths[0],
            table.ColumnWidths[table.ColumnWidths.Length - 1]);

        double maxContentHeight = fullColumnHeight;
        int[] viewportRowGroups = GetTableViewportRowGroups(tbColumn, table.Columns);
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
            if (startLine == 0 && lineCount == table.RowLineCounts[rowIndex]) return table.RowHeights[rowIndex];
            double rowLeading = table.RowLeadings[rowIndex];
            double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
            double segmentHeight = GetTableRowInitialTextHeight(tbColumn, rowIndex, table.Columns, rowLeading) + rowPadTop + rowPadBottom;
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

            // Keep first-fragment fitting consistent with the configured row height and drawing.
            return startLine == 0 ? Math.Max(segmentHeight, GetTableRowFixedHeight(tableStyle, rowIndex) ?? GetTableRowMinHeight(tableStyle, rowIndex)) : segmentHeight;
        }

        int GetColumnTableRowSegmentLineCountThatFits(int rowIndex, int startLine, double available, bool requireDefaultFirstFragment = false) {
            int remainingLines = table.RowLineCounts[rowIndex] - startLine;
            int best = 0;
            for (int candidate = 1; candidate <= remainingLines; candidate++) {
                double candidateHeight = MeasureColumnTableRowSegmentHeight(rowIndex, startLine, candidate, suppressCellObjects: false);
                double candidateClosingPadding = rowIndex == tbColumn.Rows.Count - 1 && candidate == remainingLines ? closingPadding : 0D;
                if (candidateHeight + candidateClosingPadding > available + 0.001) {
                    break;
                }

                best = candidate;
            }

            return LimitTableRowFragmentToParagraphBoundaries(table.RowLines[rowIndex], GetTableCellLayouts(tbColumn, rowIndex, table.Columns),
                startLine, best, maxContentHeight, state.Consumed > repeatHeaderHeight + 0.001D, requireDefaultFirstFragment);
        }

        bool CanSplitColumnTableRowIntoRemainingSpace(int rowIndex) =>
            rowIndex >= table.HeaderRowCount &&
            !TableRowHasViewport(tbColumn, rowIndex, table.Columns) &&
            GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex) &&
            table.RowLineCounts[rowIndex] > 1 &&
            GetColumnTableRowSegmentLineCountThatFits(rowIndex, 0, state.Remaining, requireDefaultFirstFragment: true) > 0;

        bool ShouldBreakBeforeFinalColumnTableBodyRows(int rowIndex) {
            if (viewportRowGroups[rowIndex] >= rowIndex && !StartsTableViewportRowGroup(viewportRowGroups, rowIndex)) return false;
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

        void DrawColumnTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false, double? continuedSpanHeight = null) {
            bool renderAsFooter = rowIndex >= table.FooterStartRowIndex;
            bool rowUsesBold = table.RowBold[rowIndex];
            double rowSize = table.RowSizes[rowIndex];
            double rowLeading = table.RowLeadings[rowIndex];
            bool wholeRowSegment = startLine == 0 && lineCount == table.RowLineCounts[rowIndex];
            double rowPadTop = GetTableRowMaxPaddingTop(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowPadBottom = GetTableRowMaxPaddingBottom(tbColumn, tableStyle, rowIndex, table.Columns);
            double rowHeight = continuedSpanHeight ?? MeasureColumnTableRowSegmentHeight(rowIndex, startLine, lineCount, suppressCellObjects);
            if (rowUsesBold) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            var cells = GetTableCellLayouts(tbColumn, rowIndex, table.Columns);
            double rowBottom = state.Y - rowHeight;
            if (!suppressCellObjects) table.SpanFlow.AdmitRow(rowIndex, state.Y, rowBottom);
            int bodyRowIndex = rowIndex - table.HeaderRowCount;
            bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
            bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tbColumn, rowIndex, table.Columns);
            double cornerRadius = tableCornerRadius;
            bool roundTop = cornerRadius > 0D && rowIndex == 0 && startLine == 0;
            bool roundBottom = cornerRadius > 0D && rowIndex == tbColumn.Rows.Count - 1 && startLine + lineCount >= table.RowLineCounts[rowIndex];
            bool roundedRow = roundTop || roundBottom;
            if (roundedRow) { BeginRoundedClip(sb, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom); }
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
            if (roundedRow) { EndRoundedClip(sb); }

            if (tableStyle.CellFills != null && tableStyle.CellFills.Count > 0) {
                double fillX = xTable;
                for (int fillColumn = 0; fillColumn < table.Columns; fillColumn++) {
                    if (!table.SpanFlow.Contains(rowIndex, fillColumn) && tableStyle.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
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
                        bool fillTouchesTop = cornerRadius > 0D && rowIndex == 0 && startLine == 0;
                        bool fillTouchesBottom = cornerRadius > 0D &&
                            (wholeRowSegment
                                ? rowIndex + fillCell.RowSpan >= tbColumn.Rows.Count
                                : rowIndex == tbColumn.Rows.Count - 1 && startLine + lineCount >= table.RowLineCounts[rowIndex]);
                        bool fillTouchesLeft = fillColumn == 0;
                        bool fillTouchesRight = fillColumn + span >= table.Columns;
                        bool clipCellFill = (fillTouchesTop || fillTouchesBottom) && (fillTouchesLeft || fillTouchesRight);
                        if (clipCellFill) {
                            BeginRoundedClip(sb, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, columnGap), fillHeight, cornerRadius,
                                fillTouchesTop && fillTouchesLeft,
                                fillTouchesTop && fillTouchesRight,
                                fillTouchesBottom && fillTouchesRight,
                                fillTouchesBottom && fillTouchesLeft);
                        }
                        DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(table.ColumnWidths, fillColumn, span, columnGap), fillHeight, emitGeneratedStructure);
                        if (clipCellFill) { EndRoundedClip(sb); }
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
            if (!suppressCellObjects)
                table.SpanFlow.BindStructureRow(currentPage, rowStructureElementIndex.HasValue ? currentPage!.StructElements[rowStructureElementIndex.Value] : null);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (table.SpanFlow.Contains(rowIndex, cell.Column)) continue;
                int c = cell.Column;
                xi = xTable;
                for (int xColumn = 0; xColumn < c; xColumn++) {
                    xi += table.ColumnWidths[xColumn] + columnGap;
                }

                double cellWidth = GetTableCellWidth(table.ColumnWidths, c, cell.ColumnSpan, columnGap);
                double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(table.RowHeights, rowIndex, cell.RowSpan, columnTableRowGap) : rowHeight;
                TableCellTextLayout lines = table.RowLines[rowIndex][c];
                int sourceStartLine = cell.Viewport != null || wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                int requestedLineCount = cell.Viewport != null || wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                RenderFlowTableCellContent(tbColumn, tableStyle, cell, lines, rowIndex, sourceStartLine, requestedLineCount,
                    xi, state.Y, cellWidth, cellHeight, rowSize, rowLeading, table.RowRunFontSizeScales[rowIndex],
                    rowUsesBold, renderAsHeader, wholeRowSegment, suppressCellObjects, textColor, rowStructureElementIndex);
            }

            if (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) {
                pageDirty = true;
                bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex - 1, table.Columns);
                bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tbColumn, rowIndex, table.Columns);
                bool segmentBorderRows = HasSkippedColumns(topBorderSkips, table.Columns) || HasSkippedColumns(bottomBorderSkips, table.Columns);
                if (roundedRow) {
                    if (roundTop) {
                        DrawRoundedSideStroke(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, tableStyle.BorderWidth, RoundedRectSide.Top, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    } else {
                        DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom + rowHeight, topBorderSkips, emitGeneratedStructure);
                    }
                    if (roundBottom) {
                        DrawRoundedSideStroke(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, tableStyle.BorderWidth, RoundedRectSide.Bottom, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    } else {
                        DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom, bottomBorderSkips, emitGeneratedStructure);
                    }
                    DrawRoundedSideStroke(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, tableStyle.BorderWidth, RoundedRectSide.Left, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    DrawRoundedSideStroke(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, tableStyle.BorderWidth, RoundedRectSide.Right, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                } else if (segmentBorderRows) {
                    DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom + rowHeight, topBorderSkips, emitGeneratedStructure);
                    DrawTableHorizontalLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, table.ColumnWidths, columnGap, rowBottom, bottomBorderSkips, emitGeneratedStructure);
                    DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                    DrawVLine(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable + table.Width, rowBottom + rowHeight, rowBottom, emitGeneratedStructure);
                } else {
                    DrawRoundedRowRect(sb, tableStyle.BorderColor.Value, tableStyle.BorderWidth, xTable, rowBottom, table.Width, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
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
                double roundedOuterBorder = (tableStyle.BorderColor is not null && tableStyle.BorderWidth > 0) ? tableStyle.BorderWidth : 0D;
                double borderX = xTable;
                for (int borderColumn = 0; borderColumn < table.Columns; borderColumn++) {
                    if (!table.SpanFlow.Contains(rowIndex, borderColumn) && tableStyle.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                        TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                        (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                        HasRenderableCellBorder(cellBorder)) {
                        int span = wholeRowSegment || cellBorder.HasHiddenSegments ? borderCell.ColumnSpan : 1;
                        double borderHeight = rowHeight;
                        double borderBottom = rowBottom;
                        if (wholeRowSegment) {
                            if (borderCell.RowSpan > 1) {
                                borderHeight = GetTableCellHeight(table.RowHeights, rowIndex, borderCell.RowSpan, columnTableRowGap);
                                borderBottom = state.Y - borderHeight;
                            }
                        }

                        pageDirty = true;
                        bool cellTouchesTop = cornerRadius > 0D && rowIndex == 0 && startLine == 0;
                        bool cellTouchesBottom = cornerRadius > 0D &&
                            (wholeRowSegment
                                ? rowIndex + borderCell.RowSpan >= tbColumn.Rows.Count
                                : rowIndex == tbColumn.Rows.Count - 1 && startLine + lineCount >= table.RowLineCounts[rowIndex]);
                        bool cellTouchesLeft = borderColumn == 0;
                        bool cellTouchesRight = borderColumn + span >= table.Columns;
                        bool topLeft = cellTouchesTop && cellTouchesLeft;
                        bool topRight = cellTouchesTop && cellTouchesRight;
                        bool bottomRight = cellTouchesBottom && cellTouchesRight;
                        bool bottomLeft = cellTouchesBottom && cellTouchesLeft;
                        StringBuilder borderOutput = HasPairedCellBorder(cellBorder) || table.SpanFlow.HasDeferredFills(tableStyle)
                            ? pairedBorders ??= new StringBuilder() : sb;
                        if (!cellBorder.HasHiddenSegments && (topLeft || topRight || bottomRight || bottomLeft)) {
                            DrawRoundedCellBorder(borderOutput, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight, cornerRadius, roundedOuterBorder, topLeft, topRight, bottomRight, bottomLeft, emitGeneratedStructure,
                                GetTableCellDiagonalFrame(borderCell, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight));
                        } else {
                            DrawCellBorder(borderOutput, cellBorder, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight, emitGeneratedStructure,
                                GetCellBorderSegmentLengths(table.RowHeights, rowIndex, borderCell.RowSpan, columnTableRowGap),
                                GetCellBorderSegmentLengths(table.ColumnWidths, borderColumn, span, columnGap),
                                GetTableCellDiagonalFrame(borderCell, borderX, borderBottom, GetTableCellWidth(table.ColumnWidths, borderColumn, span, columnGap), borderHeight));
                        }
                    }
                    borderX += table.ColumnWidths[borderColumn] + columnGap;
                }
            }

            if (wholeRowSegment || startLine + lineCount == table.RowLineCounts[rowIndex])
                FlushTableSpanFragments(tbColumn, tableStyle, table.SpanFlow, table.PreparedRows, table.ColumnWidths,
                    columnGap, columnTableRowGap, xTable, tableStructureElementIndex, ref pairedBorders, rowIndex,
                    tableCornerRadius, Math.Max(0D, maxContentHeight - repeatHeaderHeight), closingPadding: closingPadding);
            bool spanTailPending = GetTableSpanRemainingLineCount(table.SpanFlow, table.PreparedRows, rowIndex) > 0;
            double rowAdvance = rowHeight + (wholeRowSegment && !spanTailPending ? GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) : 0D);
            RecordTableFrameContentBottom(tbColumn, rowBottom);
            state.Y -= rowAdvance;
            state.Remaining -= rowAdvance;
            state.Consumed += rowAdvance;
        }

        void DrawColumnTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
            DrawColumnTableRowSegment(rowIndex, renderAsHeader, 0, table.RowLineCounts[rowIndex], suppressCellObjects);

        int rowIndex = state.Line;
        int rowStartLine = state.Subline;
        while (rowIndex < tbColumn.Rows.Count) {
            double maximumSpanRowHeight = Math.Max(0D, maxContentHeight -
                GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap) -
                (HasRepeatableHeader() && rowIndex >= table.HeaderRowCount ? repeatHeaderHeight : 0D));
            double spanRowRemainder = MeasureTableSpanRemainderHeight(table.SpanFlow, tableStyle,
                table.PreparedRows, table.ColumnWidths, columnGap, rowIndex, subtractAdmittedHeight: true);
            if (rowIndex == tbColumn.Rows.Count - 1 && spanRowRemainder > .001D &&
                spanRowRemainder <= maximumSpanRowHeight + .001D)
                maximumSpanRowHeight = Math.Max(0D, maximumSpanRowHeight - closingPadding);
            ReserveTableSpanRemainder(table.SpanFlow, tableStyle, table.PreparedRows, table.ColumnWidths, columnGap,
                rowIndex, maximumSpanRowHeight);
            if (table.PendingSpanTailRow == rowIndex) {
                double minimum = MeasureTableSpanRemainderHeight(table.SpanFlow, tableStyle, table.PreparedRows,
                    table.ColumnWidths, columnGap, rowIndex,
                    minimumFragmentFrameHeight: Math.Max(0D, maxContentHeight - repeatHeaderHeight -
                        (rowIndex == tbColumn.Rows.Count - 1 ? closingPadding : 0D)));
                if (HasRepeatableHeader() && AtContinuationPageTop() && repeatHeaderHeight + minimum <= state.Remaining + .001D)
                    for (int header = 0; header < table.RepeatHeaderRowCount; header++)
                        DrawColumnTableRow(header, true, suppressCellObjects: true);
                double available = state.Remaining;
                double required = MeasureTableSpanRemainderHeight(table.SpanFlow, tableStyle,
                    table.PreparedRows, table.ColumnWidths, columnGap, rowIndex);
                if (rowIndex == tbColumn.Rows.Count - 1 && required <= available + .001D)
                    available -= closingPadding;
                if (available < minimum - .001D) {
                    if (state.Consumed <= .001D)
                        throw new ArgumentException("Merged table cell content cannot fit within the available continuation frame.");
                    break;
                }
                int before = GetTableSpanRemainingLineCount(table.SpanFlow, table.PreparedRows, rowIndex);
                double height = Math.Min(available, required);
                DrawColumnTableRowSegment(rowIndex, false, table.RowLineCounts[rowIndex], 0, continuedSpanHeight: height);
                int after = GetTableSpanRemainingLineCount(table.SpanFlow, table.PreparedRows, rowIndex);
                if (after >= before)
                    throw new ArgumentException("Merged table cell content cannot fit within the available continuation frame.");
                if (after > 0) break;
                table.PendingSpanTailRow = -1;
                double gapAfterSpan = GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap);
                state.Y -= gapAfterSpan;
                state.Remaining -= gapAfterSpan;
                state.Consumed += gapAfterSpan;
                rowIndex++;
                state.Line = rowIndex;
                state.Subline = 0;
                rowStartLine = 0;
                continue;
            }
            double rowHeight = table.RowHeights[rowIndex];
            double placementHeight = GetTableViewportPlacementHeight(viewportRowGroups, table.RowHeights, rowIndex, columnTableRowGap);
            if (viewportRowGroups[rowIndex] >= rowIndex && placementHeight > maxContentHeight + 0.001D)
                throw new ArgumentException("A cell viewport's complete visible row span must fit within one page; divide it into explicit fragments before rendering.");
            double requiredRowHeight = GetTableRowFixedHeight(tableStyle, rowIndex) ?? GetTableRowMinHeight(tableStyle, rowIndex);
            if (requiredRowHeight > maxContentHeight + 0.001D)
                throw new ArgumentException("Table row height requirement exceeds the available page content height.");
            if (rowHeight > maxContentHeight + 0.001 || rowStartLine > 0) {
                if (TableRowHasViewport(tbColumn, rowIndex, table.Columns))
                    throw new ArgumentException("A row containing a cell viewport must fit within one page; divide it into explicit fragments before rendering.");
                if (!GetTableRowAllowBreakAcrossPages(tableStyle, rowIndex)) {
                    throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                }

                int totalLines = table.RowLineCounts[rowIndex];
                double minimumSegmentHeight = MeasureColumnTableRowSegmentHeight(rowIndex, rowStartLine, 1, suppressCellObjects: false);
                if (minimumSegmentHeight > maxContentHeight + 0.001D)
                    throw new ArgumentException("Table row content cannot fit within the available page content height.");
                bool repeatHeaderBeforeSegment = rowIndex >= table.HeaderRowCount &&
                    HasRepeatableHeader() &&
                    AtContinuationPageTop() &&
                    repeatHeaderHeight + minimumSegmentHeight <= state.Remaining + 0.001;
                double neededForFirstSegment = minimumSegmentHeight + (repeatHeaderBeforeSegment ? repeatHeaderHeight : 0);
                if (neededForFirstSegment > state.Remaining && state.Consumed > 0) break;
                if (neededForFirstSegment > state.Remaining && state.Consumed == 0) { state.Remaining = 0; break; }

                if (repeatHeaderBeforeSegment) {
                    for (int headerIndex = 0; headerIndex < table.RepeatHeaderRowCount; headerIndex++) {
                        DrawColumnTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                    }
                }

                int take = Math.Min(totalLines - rowStartLine, GetColumnTableRowSegmentLineCountThatFits(rowIndex, rowStartLine, state.Remaining));
                if (take == 0) break;
                DrawColumnTableRowSegment(rowIndex, renderAsHeader: rowIndex < table.HeaderRowCount && rowStartLine == 0, rowStartLine, take);
                rowStartLine += take;

                if (rowStartLine < totalLines) {
                    state.Line = rowIndex;
                    state.Subline = rowStartLine;
                    break;
                }

                if (GetTableSpanRemainingLineCount(table.SpanFlow, table.PreparedRows, rowIndex) > 0) {
                    table.PendingSpanTailRow = rowIndex;
                    state.Line = rowIndex;
                    state.Subline = totalLines;
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
                repeatHeaderHeight + placementHeight <= state.Remaining + 0.001;
            double neededForNextRow = placementHeight + (StartsTableViewportRowGroup(viewportRowGroups, rowIndex) ? 0D : GetTableRowGapAfter(rowIndex, tbColumn.Rows.Count, columnTableRowGap)) + (repeatHeaderBeforeRow ? repeatHeaderHeight : 0);
            if (rowIndex == tbColumn.Rows.Count - 1 && spanRowRemainder <= rowHeight + .001D)
                neededForNextRow += closingPadding;
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
            if (GetTableSpanRemainingLineCount(table.SpanFlow, table.PreparedRows, rowIndex) > 0) {
                table.PendingSpanTailRow = rowIndex;
                state.Line = rowIndex;
                state.Subline = table.RowLineCounts[rowIndex];
                break;
            }
            rowIndex++;
            state.Line = rowIndex;
            state.Subline = 0;
            rowStartLine = 0;
        }

        FlushTableSpanFragments(tbColumn, tableStyle, table.SpanFlow, table.PreparedRows, table.ColumnWidths,
            columnGap, columnTableRowGap, xTable, tableStructureElementIndex, ref pairedBorders,
            cornerRadius: tableCornerRadius, fullFrameHeight: Math.Max(0D, maxContentHeight - repeatHeaderHeight), closingPadding: closingPadding);
        if (pairedBorders != null) sb.Append(pairedBorders);
        if (rowIndex >= tbColumn.Rows.Count) {
            if (tableStyle.SpacingAfter > 0 && tableStyle.SpacingAfter <= state.Remaining) {
                state.Y -= tableStyle.SpacingAfter;
                state.Remaining -= tableStyle.SpacingAfter;
                state.Consumed += tableStyle.SpacingAfter;
            }
            state.Index++;
            state.Line = 0;
            state.Subline = 0;
            if (!tableStyle.ConsumesVerticalFlow && tableStartedInThisColumn) {
                state.Y = flowYBeforeTable;
                state.Remaining = flowRemainingBeforeTable;
                state.Consumed = flowConsumedBeforeTable;
            }
        } else {
            PrepareTableFrameContinuation(tbColumn, state.Line);
            return false;
        }
            return true;
        }
    }
}
