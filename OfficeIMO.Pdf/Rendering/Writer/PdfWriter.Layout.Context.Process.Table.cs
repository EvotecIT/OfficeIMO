using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderTableFlowBlock(TableBlock tb, IPdfBlock? nextBlock) {
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
            if (style.CaptionSpacingAfter < 0 || double.IsNaN(style.CaptionSpacingAfter) || double.IsInfinity(style.CaptionSpacingAfter)) {
                throw new ArgumentException("Table caption spacing after must be a non-negative finite value.");
            }
            if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
                throw new ArgumentException("Table spacing after must be a non-negative finite value.");
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
            ValidateTableCellStyleCoordinates(style, tb.Rows.Count, cols);
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
                double rowSize = GetTableRowFontSize(style, ri, headerRowCount, footerStartRowIndex, currentOpts.DefaultFontSize);
                double rowLeading = GetTableLeading(style, rowSize);
                bool rowUsesBold = GetTableRowBold(style, ri, headerRowCount, footerStartRowIndex);
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
                    TableCellTextLayout lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, rowSize, rowLeading, currentOpts);
                    rowLines[ri][cell.Column] = lines;
                    if (cell.RowSpan <= 1) {
                        maxLines = Math.Max(maxLines, lines.LineCount);
                        maxRequiredHeight = Math.Max(maxRequiredHeight, MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeading) + GetTableCellPaddingTop(style, ri, cell.Column) + GetTableCellPaddingBottom(style, ri, cell.Column));
                    }
                }
                rowLineCounts[ri] = maxLines;
                rowHeights[ri] = Math.Max(maxRequiredHeight, GetTableRowMinHeight(style, ri));
            }
            ApplyTableRowSpanHeights(tb, style, cols, rowLines, rowHeights, rowLeadings, rowGapPx);
            double xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);

            double maxContentHeight = currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
            string? captionText = string.IsNullOrWhiteSpace(style.Caption) ? null : style.Caption;
            System.Collections.Generic.List<string>? captionLines = null;
            double captionSize = style.CaptionFontSize ?? size;
            double captionLeading = captionSize * 1.25;
            double captionHeight = 0;
            if (captionText != null) {
                var captionFontForWrap = ChooseNormal(currentOpts.DefaultFont);
                captionLines = WrapSimpleTextForOptions(captionText, tableWidth, captionFontForWrap, captionSize, currentOpts).ToList();
                captionHeight = captionLines.Count * captionLeading;
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
                double nextHeight = MeasureNextBlockFirstVisualHeight(nextBlock, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
                double keepHeight = tableHeight + nextHeight;
                if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 && y < yStart - 0.001 && y - keepHeight < currentOpts.MarginBottom) {
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
                    tableStructureElementIndex = RegisterStructureContainer("Table");
                }

                return tableStructureElementIndex;
            }

            if (captionLines != null) {
                var captionFont = ChooseNormal(currentOpts.DefaultFont);
                double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                double captionAndFirstRowHeight = captionHeight + style.CaptionSpacingAfter + firstRowHeight;
                if (y < yStart - 0.001 &&
                    y - Math.Min(captionAndFirstRowHeight, maxContentHeight) < currentOpts.MarginBottom) {
                    NewPage();
                }

                int? captionMarkedContentId = RegisterTextStructureElement("Caption", EnsureTableStructureElement());
                WriteLinesInternal("F1", captionSize, captionLeading, xOrigin, tableWidth, y - GetAscenderForOptions(captionFont, captionSize, currentOpts), captionLines, style.CaptionAlign, style.CaptionColor, structureType: "Caption", markedContentId: captionMarkedContentId);
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
                repeatHeaderHeight + rowLeadings[rowIndex] + GetTableRowMaxPaddingTop(tb, style, rowIndex, cols) + GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols) <= maxContentHeight + 0.001;

            void DrawRepeatHeaders() {
                for (int headerIndex = 0; headerIndex < repeatHeaderRowCount; headerIndex++) {
                    DrawTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                }
            }

            void NewTablePage(int rowIndex) {
                NewPage();
                if (CanRepeatHeaderWithSegment(rowIndex)) {
                    DrawRepeatHeaders();
                }
            }

            void DrawTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false) {
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
                double rowHeight = wholeRowSegment ? rowHeights[rowIndex] : Math.Max(1, lineCount) * rowLeading + rowPadTop + rowPadBottom;
                double rowBottom = y - rowHeight;
                if (currentOpts.Debug?.ShowTableRowBoxes == true) { pageDirty = true; DrawRowRect(sb, new PdfColor(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight); }
                int bodyRowIndex = rowIndex - headerRowCount;
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
                var textColor = renderAsHeader ? style!.HeaderTextColor : renderAsFooter ? style!.FooterTextColor : style!.TextColor;
                int? rowStructureElementIndex = RegisterStructureContainer("TR", EnsureTableStructureElement());
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
                        double visibleContentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine, visibleLineCount, rowLeading);
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

                        WriteClippedRichParagraph(sb, paragraph, visibleLines, visibleHeights, currentOpts, firstBaseline, rowSize, rowLeading, currentPage!.Annotations, xi - TableCellClipBleed, cellBottom - TableCellClipBleed, cellWidth + (TableCellClipBleed * 2D), cellHeight + (TableCellClipBleed * 2D), xi + cellPadLeft, innerW, structureType: markedStructureType, markedContentId: markedContentId, structurePage: currentPage);
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
                        double x1 = xi + cellPadLeft;
                        double x2 = xi + cellWidth - cellPadRight;
                        double y1 = cellBottom;
                        double y2 = y;
                        currentPage!.Annotations.Add(new LinkAnnotation { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, Uri = linkUri, DestinationName = linkDestinationName, Contents = linkContents ?? cell.Text, StructElementIndex = cellLinkStructElementIndex });
                    }
                }
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
                while (startLine < totalLines) {
                    double available = y - currentOpts.MarginBottom;
                    double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
                    double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
                    double minimumRowSegmentHeight = rowLeadings[rowIndex] + rowPadTop + rowPadBottom;
                    if (available < minimumRowSegmentHeight - 0.001) {
                        NewTablePage(rowIndex);
                        available = y - currentOpts.MarginBottom;
                    }

                    int maxLinesThisPage = Math.Max(1, (int)Math.Floor((available - rowPadTop - rowPadBottom) / rowLeadings[rowIndex]));
                    int take = Math.Min(totalLines - startLine, maxLinesThisPage);
                    DrawTableRowSegment(rowIndex, renderAsHeader && startLine == 0, startLine, take);
                    startLine += take;

                    if (startLine < totalLines) {
                        NewTablePage(rowIndex);
                    }
                }
            }

            for (int rowIndex = 0; rowIndex < tb.Rows.Count; rowIndex++) {
                if (rowHeights[rowIndex] > maxContentHeight + 0.001) {
                    if (!GetTableRowAllowBreakAcrossPages(style, rowIndex)) {
                        throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                    }

                    DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    continue;
                }

                if (ShouldBreakBefore(rowHeights[rowIndex])) {
                    NewPage();
                    if (hasRepeatableHeader && rowIndex >= headerRowCount && repeatHeaderHeight + rowHeights[rowIndex] <= maxContentHeight + 0.001) {
                        DrawRepeatHeaders();
                    }
                }

                DrawTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
            }

            y -= style.SpacingAfter;
        }

    }
}
