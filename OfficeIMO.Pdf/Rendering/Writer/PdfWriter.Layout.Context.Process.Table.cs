using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderDeferredTableFlowBlock(DeferredTableBlock deferredTable, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfTableStyle style = deferredTable.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            foreach (DeferredTableBatch batch in deferredTable.CreateBatches(style)) {
                RenderTableFlowBlock(
                    batch.Table,
                    batch.IsLast ? nextBlock : null,
                    blockList,
                    blockIndex,
                    skipInitialHeaderRows: !batch.IsFirst,
                    bodyRowOffset: batch.BodyRowOffset);
            }
        }

        private void RenderTableFlowBlock(TableBlock tb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex, bool skipInitialHeaderRows = false, int bodyRowOffset = 0) {
            PdfTableStyle style = tb.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            int cols = GetTableColumnCount(tb);
            if (cols == 0) return;
            double padLeft = GetTableCellPaddingLeft(style);
            double padRight = GetTableCellPaddingRight(style);
            double padTop = GetTableCellPaddingTop(style);
            double padBottom = GetTableCellPaddingBottom(style);
            double cellSpacing = GetTableCellSpacing(style);
            double colGapPx = cellSpacing;
            double rowGapPx = cellSpacing;
            double size = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
            if (!IsValidPdfAlign(tb.Align)) {
                throw new ArgumentException("Table alignment must be Left, Center, or Right.");
            }
            if (style.Alignments != null) {
                foreach (var alignment in style.Alignments) {
                    if (!IsValidPdfColumnAlign(alignment)) {
                        throw new ArgumentException("Table column alignments must be Left, Center, or Right.");
                    }
                }
            }
            if (style.VerticalAlignments != null) {
                foreach (var verticalAlignment in style.VerticalAlignments) {
                    if (!IsValidPdfCellVerticalAlign(verticalAlignment)) {
                        throw new ArgumentException("Table vertical alignments must be defined PDF cell vertical alignment values.");
                    }
                }
            }
            if (!IsValidPdfAlign(style.CaptionAlign)) {
                throw new ArgumentException("Table caption alignment must be Left, Center, or Right.");
            }
            if (style.BorderWidth < 0 || double.IsNaN(style.BorderWidth) || double.IsInfinity(style.BorderWidth)) {
                throw new ArgumentException("Table border width must be a non-negative finite value.");
            }
            if (style.RowSeparatorWidth < 0 || double.IsNaN(style.RowSeparatorWidth) || double.IsInfinity(style.RowSeparatorWidth)) {
                throw new ArgumentException("Table row separator width must be a non-negative finite value.");
            }
            if (style.HeaderSeparatorWidth < 0 || double.IsNaN(style.HeaderSeparatorWidth) || double.IsInfinity(style.HeaderSeparatorWidth)) {
                throw new ArgumentException("Table header separator width must be a non-negative finite value.");
            }
            if (style.FooterSeparatorWidth < 0 || double.IsNaN(style.FooterSeparatorWidth) || double.IsInfinity(style.FooterSeparatorWidth)) {
                throw new ArgumentException("Table footer separator width must be a non-negative finite value.");
            }
            if (style.CellPaddingX < 0 || double.IsNaN(style.CellPaddingX) || double.IsInfinity(style.CellPaddingX)) {
                throw new ArgumentException("Table horizontal cell padding must be a non-negative finite value.");
            }
            if (style.CellPaddingY < 0 || double.IsNaN(style.CellPaddingY) || double.IsInfinity(style.CellPaddingY)) {
                throw new ArgumentException("Table vertical cell padding must be a non-negative finite value.");
            }
            if (style.MinRowHeight < 0 || double.IsNaN(style.MinRowHeight) || double.IsInfinity(style.MinRowHeight)) {
                throw new ArgumentException("Table minimum row height must be a non-negative finite value.");
            }
            if (style.RowMinHeights != null) {
                foreach (double? rowMinHeight in style.RowMinHeights) {
                    if (rowMinHeight.HasValue && (rowMinHeight.Value < 0 || double.IsNaN(rowMinHeight.Value) || double.IsInfinity(rowMinHeight.Value))) {
                        throw new ArgumentException("Table row minimum heights must be non-negative finite values.");
                    }
                }
            }
            if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
                throw new ArgumentException("Table spacing before must be a non-negative finite value.");
            }
            if (style.Caption != null && string.IsNullOrWhiteSpace(style.Caption)) {
                throw new ArgumentException("Table caption cannot be empty or whitespace.");
            }
            if (style.CaptionFontSize.HasValue && (style.CaptionFontSize.Value <= 0 || double.IsNaN(style.CaptionFontSize.Value) || double.IsInfinity(style.CaptionFontSize.Value))) {
                throw new ArgumentException("Table caption font size must be a positive finite value.");
            }
            if (style.MinimumShrinkFontSize.HasValue && (style.MinimumShrinkFontSize.Value <= 0 || double.IsNaN(style.MinimumShrinkFontSize.Value) || double.IsInfinity(style.MinimumShrinkFontSize.Value))) {
                throw new ArgumentException("Table minimum shrink font size must be a positive finite value.");
            }
            if (style.CaptionSpacingAfter < 0 || double.IsNaN(style.CaptionSpacingAfter) || double.IsInfinity(style.CaptionSpacingAfter)) {
                throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
            }
            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Table spacing after must be a non-negative finite value.");
            }
            if (style.PageContinuationSpacingBefore < 0 || double.IsNaN(style.PageContinuationSpacingBefore) || double.IsInfinity(style.PageContinuationSpacingBefore)) {
                throw new ArgumentException("Table page continuation spacing before must be a non-negative finite value.");
            }
            if (double.IsNaN(style.RowBaselineOffset) || double.IsInfinity(style.RowBaselineOffset)) {
                throw new ArgumentException("Table row baseline offset must be a finite value.");
            }
            if (style.CellFills != null) {
                foreach (var cellFill in style.CellFills) {
                    if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                        throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                    }
                }
            }
            if (style.CellBorders != null) {
                foreach (var cellBorder in style.CellBorders) {
                    if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                        throw new ArgumentException("Table cell border coordinates cannot be negative.");
                    }
                    if (cellBorder.Value == null || cellBorder.Value.Width < 0 || double.IsNaN(cellBorder.Value.Width) || double.IsInfinity(cellBorder.Value.Width)) {
                        throw new ArgumentException("Table cell border widths must be non-negative finite values.");
                    }
                }
            }
            if (style.HeaderRowCount < 0) {
                throw new ArgumentException("Table header row count cannot be negative.");
            }
            if (style.FooterRowCount < 0) {
                throw new ArgumentException("Table footer row count cannot be negative.");
            }

            ValidateTableRoleRowCounts(style, tb.Rows.Count);
            int headerRowCount = style.HeaderRowCount;
            int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
            int footerRowCount = style.FooterRowCount;
            int footerStartRowIndex = tb.Rows.Count - footerRowCount;
            ValidateTableCellStyleCoordinates(style, tb, cols);
            ValidateTableColumnStyleBounds(style, cols);
            ValidateTableRowStyleBounds(style, tb.Rows.Count);
            ValidateTableRowSpansWithinRoleBoundaries(tb, cols, headerRowCount, footerStartRowIndex);
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PreparedTableColumns preparedColumns = PrepareTableColumns(tb, style, contentWidth, size, headerRowCount, footerStartRowIndex);
            double tableWidth = preparedColumns.TableWidth;
            double[] colPixel = preparedColumns.ColumnWidths;
            ValidateTableCellTextWidths(tb, style, cols, colPixel, colGapPx);

            var rowLines = new TableCellTextLayout[tb.Rows.Count][];
            var rowLineCounts = new int[tb.Rows.Count];
            var rowHeights = new double[tb.Rows.Count];
            var rowLeadings = new double[tb.Rows.Count];
            var rowSizes = new double[tb.Rows.Count];
            var rowBold = new bool[tb.Rows.Count];
            for (int ri = 0; ri < tb.Rows.Count; ri++) {
                double originalRowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
                double rowSize = ResolveTableRowShrinkFontSize(tb, style, ri, cols, colPixel, colGapPx, originalRowSize, rowUsesBold, currentOpts);
                double runFontSizeScale = GetTableRunFontSizeScale(tb, style, ri, cols, colPixel, colGapPx, originalRowSize, rowSize, rowUsesBold, currentOpts);
                double rowLeading = GetTableLeading(style, rowSize);
                rowSizes[ri] = rowSize;
                rowLeadings[ri] = rowLeading;
                rowBold[ri] = rowUsesBold;
                rowLines[ri] = new TableCellTextLayout[cols];
                int maxLines = 1;
                double maxRequiredHeight = rowLeading + GetTableRowMaxPaddingTop(tb, style, ri, cols) + GetTableRowMaxPaddingBottom(tb, style, ri, cols);
                for (int ci = 0; ci < cols; ci++) {
                    rowLines[ri][ci] = new TableCellTextLayout(new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> { new() }, new System.Collections.Generic.List<double> { rowLeading });
                }

                var cells = GetTableCellLayouts(tb, ri, cols);
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                    double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, colGapPx);
                    double innerWidth = Math.Max(1, cellWidth - GetTableCellPaddingLeft(style, ri, cell.Column) - GetTableCellPaddingRight(style, ri, cell.Column));
                    TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading, currentOpts, runFontSizeScale, style.MinimumShrinkFontSize ?? 6D);
                    rowLines[ri][cell.Column] = lines;
                    if (cell.RowSpan <= 1) {
                        maxLines = Math.Max(maxLines, lines.LineCount);
                        maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading, innerWidth) + GetTableCellPaddingTop(style, ri, cell.Column) + GetTableCellPaddingBottom(style, ri, cell.Column));
                    }
                }
                rowLineCounts[ri] = maxLines;
                rowHeights[ri] = ResolveTableRowHeight(style, ri, maxRequiredHeight);
            }
            ApplyTableRowSpanHeights(tb, style, cols, colPixel, rowLines, rowHeights, rowLeadings, colGapPx, rowGapPx);
            double xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);

            double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            string? captionText = string.IsNullOrWhiteSpace(style.Caption) ? null : style.Caption;
            System.Collections.Generic.IReadOnlyList<TextRun>? captionRuns = null;
            System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>? captionLines = null;
            System.Collections.Generic.List<double>? captionLineHeights = null;
            double captionSize = style.CaptionFontSize ?? size;
            double captionLeading = captionSize * 1.25;
            double captionHeight = 0;
            if (captionText != null) {
                var captionFontForWrap = ChooseNormal(currentOpts.DefaultFont);
                captionRuns = new[] { TextRun.Normal(captionText, style.CaptionColor, captionSize) };
                var captionWrap = WrapRichRunsCore(captionRuns, tableWidth, captionSize, captionFontForWrap, captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                captionLines = captionWrap.Lines;
                captionLineHeights = captionWrap.LineHeights;
                captionHeight = MeasureRichLinesHeight(captionLineHeights, captionLines.Count, captionLeading);
                double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                if (captionHeight + style.CaptionSpacingAfter + firstRowHeight > maxContentHeight + 0.001) {
                    throw new ArgumentException("Table caption and first row exceed the available page content height.");
                }
            }

            double tableContentHeight = (captionLines == null ? 0 : captionHeight + style.CaptionSpacingAfter) + GetTableRowsHeight(rowHeights, 0, rowHeights.Length, rowGapPx);
            double tableSpacingBefore = y < yStart - 0.001 ? style.SpacingBefore : 0D;
            if (style.KeepTogether) {
                double keepHeight = tableSpacingBefore + tableContentHeight + style.SpacingAfter;
                if (keepHeight > maxContentHeight + 0.001) {
                    throw new ArgumentException("Table height exceeds the available page content height.");
                }

                if (y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    tableSpacingBefore = 0D;
                }
            }

            if (style.KeepWithNext && nextBlock != null) {
                double tableHeight = tableSpacingBefore + tableContentHeight + style.SpacingAfter;
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                double keepHeight = tableHeight + nextHeight;
                if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
                    NewPage();
                    tableSpacingBefore = 0D;
                }
            }

            int minimumFirstPageBodyRows = Math.Min(
                style.MinimumBodyRowsOnFirstPage,
                Math.Max(0, footerStartRowIndex - headerRowCount));
            if (!skipInitialHeaderRows && minimumFirstPageBodyRows > 0 && y < yStart - 0.001) {
                int firstPageRowCount = headerRowCount + minimumFirstPageBodyRows;
                double firstPageGroupHeight =
                    tableSpacingBefore +
                    (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter) +
                    GetTableRowsHeight(rowHeights, 0, firstPageRowCount, rowGapPx);
                if (firstPageGroupHeight <= maxContentHeight + 0.001 &&
                    y - firstPageGroupHeight < currentOpts.MarginBottom) {
                    NewPage();
                    tableSpacingBefore = 0D;
                }
            }

            if (tableSpacingBefore > 0) {
                if (y < yStart - 0.001 && y - tableSpacingBefore < currentOpts.MarginBottom) {
                    NewPage();
                    tableSpacingBefore = 0D;
                }

                y -= tableSpacingBefore;
            }

            int? tableStructureElementIndex = null;
            LayoutResult.Page? tableStructurePage = null;
            int? EnsureTableStructureElement() {
                if (!emitGeneratedStructure || currentPage == null) {
                    return null;
                }

                if (!ReferenceEquals(tableStructurePage, currentPage)) {
                    tableStructurePage = currentPage;
                    tableStructureElementIndex = RegisterStructureContainer("Table", alternativeText: style.AlternativeText);
                }

                return tableStructureElementIndex;
            }

            if (captionRuns != null && captionLines != null && captionLineHeights != null) {
                var captionFont = ChooseNormal(currentOpts.DefaultFont);
                double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                double captionAndFirstRowHeight = captionHeight + style.CaptionSpacingAfter + firstRowHeight;
                if (y < yStart - 0.001 &&
                    y - Math.Min(captionAndFirstRowHeight, maxContentHeight) < currentOpts.MarginBottom) {
                    NewPage();
                }

                int? captionMarkedContentId = RegisterTextStructureElement("Caption", EnsureTableStructureElement());
                MarkRichFonts(captionRuns);
                WriteRichParagraph(sb, new RichParagraphBlock(captionRuns, style.CaptionAlign, style.CaptionColor), captionLines, captionLineHeights, currentOpts, FirstTextBaselineFromTop(captionFont, captionSize, y), captionSize, captionLeading, currentPage!.Annotations, xOrigin, tableWidth, structureType: "Caption", markedContentId: captionMarkedContentId, structurePage: currentPage);
                y -= captionHeight + style.CaptionSpacingAfter;
            }

            if (TableUsesBold(style, tb.Rows.Count, headerRowCount, footerStartRowIndex)) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            bool hasRepeatableHeader = repeatHeaderRowCount > 0 && tb.Rows.Count > headerRowCount;
            double repeatHeaderHeight = 0;
            for (int i = 0; i < repeatHeaderRowCount; i++) {
                repeatHeaderHeight += rowHeights[i] + GetTableRowGapAfter(i, tb.Rows.Count, rowGapPx);
            }

            bool ShouldBreakBefore(double rowHeight) =>
                y < yStart - 0.001 &&
                y - rowHeight < currentOpts.MarginBottom &&
                rowHeight <= maxContentHeight;

            bool CanRepeatHeaderWithSegment(int rowIndex) =>
                hasRepeatableHeader &&
                rowIndex >= headerRowCount &&
                repeatHeaderHeight + rowLeadings[rowIndex] + GetTableRowMaxPaddingTop(tb, style, rowIndex, cols) + GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols) <= y - currentOpts.MarginBottom + 0.001;

            void ApplyTablePageContinuationSpacing(double requiredFirstSegmentHeight) {
                double spacing = style.PageContinuationSpacingBefore;
                if (spacing <= 0D || y < yStart - 0.001) {
                    return;
                }

                double available = y - currentOpts.MarginBottom;
                double spacingThatFits = Math.Min(spacing, Math.Max(0D, available - requiredFirstSegmentHeight));
                if (spacingThatFits > 0.001D) {
                    y -= spacingThatFits;
                }
            }

            double GetTableContinuationRequiredHeight(int rowIndex) {
                double rowRequiredHeight = rowHeights[rowIndex];
                if (hasRepeatableHeader && rowIndex >= headerRowCount) {
                    rowRequiredHeight += repeatHeaderHeight;
                }

                return rowRequiredHeight;
            }

            void DrawRepeatHeaders() {
                for (int headerIndex = 0; headerIndex < repeatHeaderRowCount; headerIndex++) {
                    DrawTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                }
            }

            void NewTablePage(int rowIndex) {
                NewPage();
                ApplyTablePageContinuationSpacing(GetTableContinuationRequiredHeight(rowIndex));
                if (CanRepeatHeaderWithSegment(rowIndex)) {
                    DrawRepeatHeaders();
                }
            }

            double MeasureTableRowSegmentHeight(int rowIndex, int startLine, int lineCount, bool suppressCellObjects) {
                double rowLeading = rowLeadings[rowIndex];
                double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                double segmentHeight = rowLeading + rowPadTop + rowPadBottom;
                var cells = GetTableCellLayouts(tb, rowIndex, cols);
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    double cellWidth = GetTableCellWidth(colPixel, cell.Column, cell.ColumnSpan, colGapPx);
                    double cellPadLeft = GetTableCellPaddingLeft(style, rowIndex, cell.Column);
                    double cellPadRight = GetTableCellPaddingRight(style, rowIndex, cell.Column);
                    double innerW = cellWidth - cellPadLeft - cellPadRight;
                    TableCellTextLayout lines = rowLines[rowIndex][cell.Column];
                    int sourceStartLine = startLine;
                    int visibleLineCount = Math.Max(0, Math.Min(lineCount, lines.LineCount - sourceStartLine));
                    bool includeObjects = !suppressCellObjects && sourceStartLine == 0;
                    double cellContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW, includeObjects) +
                        GetTableCellPaddingTop(style, rowIndex, cell.Column) +
                        GetTableCellPaddingBottom(style, rowIndex, cell.Column);
                    segmentHeight = Math.Max(segmentHeight, cellContentHeight);
                }

                return segmentHeight;
            }

            int GetTableRowSegmentLineCountThatFits(int rowIndex, int startLine, double available) {
                int remaining = rowLineCounts[rowIndex] - startLine;
                int best = 0;
                for (int candidate = 1; candidate <= remaining; candidate++) {
                    double candidateHeight = MeasureTableRowSegmentHeight(rowIndex, startLine, candidate, suppressCellObjects: false);
                    if (candidateHeight > available + 0.001) {
                        break;
                    }

                    best = candidate;
                }

                return Math.Max(1, best);
            }

            bool CanSplitTableRowIntoRemainingSpace(int rowIndex) =>
                rowIndex >= headerRowCount &&
                GetTableRowAllowBreakAcrossPages(style, rowIndex) &&
                rowLineCounts[rowIndex] > 1 &&
                MeasureTableRowSegmentHeight(rowIndex, 0, Math.Min(2, rowLineCounts[rowIndex]), suppressCellObjects: false) <= y - currentOpts.MarginBottom + 0.001;

            bool ShouldBreakBeforeFinalBodyRows(int rowIndex) {
                int minimumBodyRows = Math.Min(style.MinimumBodyRowsOnLastPage, Math.Max(0, footerStartRowIndex - headerRowCount));
                if (minimumBodyRows <= 0 || footerStartRowIndex - rowIndex != minimumBodyRows) {
                    return false;
                }

                double currentRowHeight = rowHeights[rowIndex] + GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                double finalGroupHeight = GetTableRowsHeight(rowHeights, rowIndex, rowHeights.Length, rowGapPx);
                return ShouldBreakBeforeFinalTableBodyRows(
                    rowIndex,
                    headerRowCount,
                    footerStartRowIndex,
                    minimumBodyRows,
                    currentRowHeight,
                    finalGroupHeight,
                    y - currentOpts.MarginBottom,
                    hasRepeatableHeader ? repeatHeaderHeight : 0D,
                    maxContentHeight,
                    y < yStart - 0.001);
            }

            void DrawTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false, PageStructElement? existingRowStructureElement = null) {
                bool renderAsFooter = rowIndex >= footerStartRowIndex;
                bool rowUsesBold = rowBold[rowIndex];
                double rowSize = rowSizes[rowIndex];
                double rowLeading = rowLeadings[rowIndex];
                if (rowUsesBold) {
                    currentPage!.UsedBold = true;
                    usedBold = true;
                }

                var cells = GetTableCellLayouts(tb, rowIndex, cols);
                bool wholeRowSegment = startLine == 0 && lineCount == rowLineCounts[rowIndex];
                double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                double rowHeight = wholeRowSegment ? rowHeights[rowIndex] : MeasureTableRowSegmentHeight(rowIndex, startLine, lineCount, suppressCellObjects);
                double rowBottom = y - rowHeight;
                if (currentOpts.Debug?.ShowTableRowBoxes == true) { pageDirty = true; DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight); }
                int bodyRowIndex = bodyRowOffset + rowIndex - headerRowCount;
                bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
                bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tb, rowIndex, cols);
                if (style?.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, style.HeaderFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); } else if (style?.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, style.FooterFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); } else if (!renderAsHeader && !renderAsFooter && style?.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, style.RowStripeFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
                if (!renderAsHeader && !renderAsFooter && style?.BodyColumnFills != null) {
                    bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tb, rowIndex, cols);
                    double fillX = xOrigin;
                    for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                        PdfColor? fill = fillColumn < style.BodyColumnFills.Count ? style.BodyColumnFills[fillColumn] : null;
                        if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                            pageDirty = true;
                            DrawRowFill(sb, fill.Value, fillX, rowBottom, colPixel[fillColumn], rowHeight, emitGeneratedStructure);
                        }
                        fillX += colPixel[fillColumn] + colGapPx;
                    }
                }
                if (style?.CellFills != null && style.CellFills.Count > 0) {
                    double fillX = xOrigin;
                    for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                        if (style.CellFills.TryGetValue((rowIndex, fillColumn), out PdfColor fill) &&
                            TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                            (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                            pageDirty = true;
                            int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                            double fillHeight = rowHeight;
                            double fillBottom = rowBottom;
                            if (wholeRowSegment) {
                                if (fillCell.RowSpan > 1) {
                                    fillHeight = GetTableCellHeight(rowHeights, rowIndex, fillCell.RowSpan, rowGapPx);
                                    fillBottom = y - fillHeight;
                                }
                            }

                            DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(colPixel, fillColumn, span, colGapPx), fillHeight, emitGeneratedStructure);
                        }
                        fillX += colPixel[fillColumn] + colGapPx;
                    }
                }
                if (style != null && DrawTableCellDataBars(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                    pageDirty = true;
                }
                if (style != null && DrawTableCellIcons(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                    pageDirty = true;
                }
                if (currentOpts.Debug?.ShowTableBaselines == true) {
                    double x1 = xOrigin;
                    double x2 = xOrigin + tableWidth;
                    double baselineYDbg = y - padTop - GetAscenderForOptions(GetTableRowFont(currentOpts, rowUsesBold), rowSize, currentOpts);
                    pageDirty = true;
                    DrawHLine(sb, new PdfColor(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
                }
                double xi = xOrigin;
                double yRect = rowBottom;
                double rowWidth = tableWidth;
                double hRect = rowHeight;
                if (style?.BorderColor is not null && style.BorderWidth > 0) {
                    pageDirty = true;
                    bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols);
                    bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex, cols);
                    bool segmentBorderRows = HasSkippedColumns(topBorderSkips, cols) || HasSkippedColumns(bottomBorderSkips, cols);
                    if (segmentBorderRows) {
                        DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect + hRect, topBorderSkips, emitGeneratedStructure);
                        DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, colPixel, colGapPx, yRect, bottomBorderSkips, emitGeneratedStructure);
                        DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect + hRect, yRect, emitGeneratedStructure);
                        DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xOrigin + tableWidth, yRect + hRect, yRect, emitGeneratedStructure);
                    } else {
                        DrawRowRect(sb, style.BorderColor.Value, style.BorderWidth, xOrigin, yRect, rowWidth, hRect, emitGeneratedStructure);
                    }

                    double xi2 = xOrigin;
                    double yTop = yRect + hRect;
                    double yBottom = yRect;
                    for (int c = 0; c < cols - 1; c++) {
                        xi2 += colPixel[c];
                        if (IsTableBoundaryInsideSpannedCell(tb, rowIndex, c, cols)) {
                            xi2 += colGapPx;
                            continue;
                        }

                        if (currentOpts.Debug?.ShowTableColumnGuides == true)
                            DrawVLine(sb, new PdfColor(0, 0, 1), Math.Max(0.3, style.BorderWidth), xi2, yTop, yBottom);
                        else
                            DrawVLine(sb, style.BorderColor.Value, style.BorderWidth, xi2, yTop, yBottom, emitGeneratedStructure);
                        xi2 += colGapPx;
                    }
                }
                if (style != null && renderAsFooter && rowIndex == footerStartRowIndex) {
                    PdfColor? footerSeparatorColor = style.FooterSeparatorColor ?? style.RowSeparatorColor;
                    double footerSeparatorWidth = style.FooterSeparatorWidth > 0 ? style.FooterSeparatorWidth : style.RowSeparatorWidth;
                    if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                        pageDirty = true;
                        DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xOrigin, colPixel, colGapPx, y, GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols), emitGeneratedStructure);
                    }
                }
                PdfColor? separatorColor = renderAsHeader && style?.HeaderSeparatorColor is not null ? style.HeaderSeparatorColor : style?.RowSeparatorColor;
                double separatorWidth = renderAsHeader && style?.HeaderSeparatorWidth > 0 ? style.HeaderSeparatorWidth : style?.RowSeparatorWidth ?? 0;
                if (separatorColor is not null && separatorWidth > 0) {
                    pageDirty = true;
                    DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xOrigin, colPixel, colGapPx, rowBottom, GetRowSpanBoundarySkipColumns(tb, rowIndex, cols), emitGeneratedStructure);
                }
                var textColor = renderAsHeader ? style!.HeaderTextColor : renderAsFooter ? style!.FooterTextColor : style!.TextColor;
                int? rowStructureElementIndex = null;
                PageStructElement? rowStructureElement = existingRowStructureElement;
                if (rowStructureElement == null) {
                    rowStructureElementIndex = RegisterStructureContainer("TR", EnsureTableStructureElement());
                    if (rowStructureElementIndex.HasValue) {
                        rowStructureElement = currentPage!.StructElements[rowStructureElementIndex.Value];
                    }
                }
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    int c = cell.Column;
                    xi = xOrigin;
                    for (int xColumn = 0; xColumn < c; xColumn++) {
                        xi += colPixel[xColumn] + colGapPx;
                    }

                    double cellWidth = GetTableCellWidth(colPixel, c, cell.ColumnSpan, colGapPx);
                    double cellPadLeft = GetTableCellPaddingLeft(style, rowIndex, c);
                    double cellPadRight = GetTableCellPaddingRight(style, rowIndex, c);
                    double cellPadTop = GetTableCellPaddingTop(style, rowIndex, c);
                    double cellPadBottom = GetTableCellPaddingBottom(style, rowIndex, c);
                    double innerW = cellWidth - cellPadLeft - cellPadRight;
                    double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGapPx) : rowHeight;
                    double cellBottom = y - cellHeight;
                    PdfColumnAlign align = GetTableCellAlignment(style, rowIndex, c, cell.Text);
                    PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(style, rowIndex, c);

                    var cellFont = GetTableRowFont(currentOpts, rowUsesBold);
                    TableCellTextLayout lines = rowLines[rowIndex][c];
                    int sourceStartLine = wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                    int requestedLineCount = wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                    int visibleLineCount = Math.Max(0, Math.Min(requestedLineCount, lines.LineCount - sourceStartLine));
                    double verticalOffset = 0;
                    double visibleTextHeight = 0D;
                    if (visibleLineCount > 0) {
                        double availableTextHeight = Math.Max(0, cellHeight - cellPadTop - cellPadBottom);
                        visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleLineCount, rowLeading);
                        double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading, innerW);
                        double unusedTextHeight = Math.Max(0, availableTextHeight - visibleContentHeight);
                        if (verticalAlign == PdfCellVerticalAlign.Middle) verticalOffset = unusedTextHeight / 2;
                        else if (verticalAlign == PdfCellVerticalAlign.Bottom) verticalOffset = unusedTextHeight;
                    }

                    double firstBaseline = y - cellPadTop - verticalOffset - GetAscenderForOptions(cellFont, rowSize, currentOpts) + style.RowBaselineOffset;

                    pageDirty = true;
                    if (cell.Runs.Any(run => run.Bold || rowUsesBold)) { currentPage!.UsedBold = true; usedBold = true; }
                    if (cell.Runs.Any(run => run.Italic)) { currentPage!.UsedItalic = true; usedItalic = true; }
                    if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) { currentPage!.UsedBoldItalic = true; usedBoldItalic = true; }
                    MarkRichFonts(cell.Runs);
                    string? linkUri = cell.LinkUri;
                    string? linkDestinationName = cell.LinkDestinationName;
                    string? linkContents = cell.LinkContents;
                    if (tb.Links.TryGetValue((rowIndex, c), out var uri)) {
                        linkUri = uri;
                        linkDestinationName = null;
                        linkContents = cell.Text;
                    }

                    if (sourceStartLine == 0) {
                        AddTableCellNamedDestinationName(cell.NamedDestinationName, y);
                    }

                    int? cellLinkStructElementIndex = null;
                    if (visibleLineCount > 0) {
                        var visibleLines = SliceTableCellLines(lines, sourceStartLine, visibleLineCount);
                        visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
                        var visibleHeights = SliceTableCellLineHeights(lines, sourceStartLine, visibleLineCount, rowLeading);
                        var visibleAlignments = SliceTableCellLineAlignments(lines, sourceStartLine, visibleLineCount);
                        var visibleXOffsets = SliceTableCellLineXOffsets(lines, sourceStartLine, visibleLineCount);
                        var visibleWidths = SliceTableCellLineWidths(lines, sourceStartLine, visibleLineCount, innerW);
                        var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
                        string structureType = renderAsHeader ? "TH" : "TD";
                        int tableColumnSpan = cell.ColumnSpan > 1 ? cell.ColumnSpan : 1;
                        int tableRowSpan = wholeRowSegment && cell.RowSpan > 1 ? cell.RowSpan : 1;
                        bool cellHasLinkTarget = HasCellLinkTarget(linkUri, linkDestinationName);
                        int? markedContentId;
                        string markedStructureType = structureType;
                        if (cellHasLinkTarget && emitGeneratedStructure && currentPage != null) {
                            PageStructElement? cellElement = rowStructureElement == null
                                ? null
                                : RegisterStructureContainer(structureType, rowStructureElement, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                            int? cellElementIndex = cellElement == null
                                ? RegisterStructureContainer(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan)
                                : null;
                            markedStructureType = "Link";
                            markedContentId = cellElement == null
                                ? RegisterTextStructureElement(markedStructureType, cellElementIndex)
                                : RegisterTextStructureElement(markedStructureType, cellElement);
                            cellLinkStructElementIndex = FindStructElementIndex(currentPage, markedContentId, markedStructureType);
                        } else {
                            markedContentId = rowStructureElement == null
                                ? RegisterTextStructureElement(structureType, rowStructureElementIndex, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan)
                                : RegisterTextStructureElement(structureType, rowStructureElement, renderAsHeader ? "Column" : string.Empty, tableColumnSpan, tableRowSpan);
                        }

                        WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + (TableCellClipBleed * 2D), cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage, lineAlignments: visibleAlignments, lineXOffsets: visibleXOffsets, lineWidths: visibleWidths);
                    }
                    if (!suppressCellObjects && (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) && sourceStartLine == 0) {
                        if (CanRenderTableCellCheckBoxInline(cell, lines, sourceStartLine, visibleLineCount)) {
                            RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[sourceStartLine], xi + cellPadLeft, innerW, firstBaseline);
                        } else {
                            double formFieldTop = y - cellPadTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : visibleTextHeight + TableCellCheckBoxGap);
                            RenderTableCellObjects(currentPage!, cell, align, xi + cellPadLeft, innerW, formFieldTop);
                        }
                    }

                    if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                        double x1 = xi + cellPadLeft - TableCellClipBleed;
                        double x2 = xi + cellWidth - cellPadRight + TableCellClipBleed;
                        double linkCellHeight = sourceStartLine == 0 && cell.RowSpan > 1
                            ? GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGapPx)
                            : cellHeight;
                        double y1 = y - linkCellHeight - TableCellClipBleed;
                        double y2 = y + TableCellClipBleed;
                        currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text, StructElementIndex = cellLinkStructElementIndex });
                    }
                }
                if (style?.CellBorders != null && style.CellBorders.Count > 0) {
                    double borderX = xOrigin;
                    for (int borderColumn = 0; borderColumn < cols; borderColumn++) {
                        if (style.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                            TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                            (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                            HasRenderableCellBorder(cellBorder)) {
                            int span = wholeRowSegment ? borderCell.ColumnSpan : 1;
                            double borderHeight = hRect;
                            double borderBottom = yRect;
                            if (wholeRowSegment) {
                                if (borderCell.RowSpan > 1) {
                                    borderHeight = GetTableCellHeight(rowHeights, rowIndex, borderCell.RowSpan, rowGapPx);
                                    borderBottom = y - borderHeight;
                                }
                            }

                            pageDirty = true;
                            DrawCellBorder(sb, cellBorder, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight, emitGeneratedStructure);
                        }
                        borderX += colPixel[borderColumn] + colGapPx;
                    }
                }
                y -= rowHeight;
                if (wholeRowSegment) {
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                }
            }

            void DrawTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
                DrawTableRowSegment(rowIndex, renderAsHeader, 0, rowLineCounts[rowIndex], suppressCellObjects);

            void DrawSplitTableRow(int rowIndex, bool renderAsHeader) {
                int startLine = 0;
                int totalLines = rowLineCounts[rowIndex];
                PageStructElement? rowStructureElement = null;
                while (startLine < totalLines) {
                    double available = y - currentOpts.MarginBottom;
                    double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                    double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                    double minimumRowSegmentHeight = rowLeadings[rowIndex] + rowPadTop + rowPadBottom;
                    if (available < minimumRowSegmentHeight - 0.001) {
                        NewTablePage(rowIndex);
                        available = y - currentOpts.MarginBottom;
                    }

                    int take = Math.Min(totalLines - startLine, GetTableRowSegmentLineCountThatFits(rowIndex, startLine, available));
                    DrawTableRowSegment(rowIndex, renderAsHeader && startLine == 0, startLine, take, existingRowStructureElement: rowStructureElement);
                    if (rowStructureElement == null && emitGeneratedStructure && currentPage != null) {
                        rowStructureElement = currentPage.StructElements.LastOrDefault(element => element.StructureType == "TR");
                    }
                    startLine += take;

                    if (startLine < totalLines) {
                        NewTablePage(rowIndex);
                    }
                }
            }

            int firstRowIndex = skipInitialHeaderRows ? headerRowCount : 0;
            for (int rowIndex = firstRowIndex; rowIndex < tb.Rows.Count; rowIndex++) {
                if (rowHeights[rowIndex] > maxContentHeight + 0.001) {
                    if (!GetTableRowAllowBreakAcrossPages(style, rowIndex)) {
                        throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                    }

                    DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    continue;
                }

                if (ShouldBreakBefore(rowHeights[rowIndex])) {
                    if (CanSplitTableRowIntoRemainingSpace(rowIndex)) {
                        DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                        y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                        continue;
                    }

                    NewTablePage(rowIndex);
                } else if (ShouldBreakBeforeFinalBodyRows(rowIndex)) {
                    NewTablePage(rowIndex);
                }

                DrawTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
            }

            y -= style.SpacingAfter;
        }

    }
}
