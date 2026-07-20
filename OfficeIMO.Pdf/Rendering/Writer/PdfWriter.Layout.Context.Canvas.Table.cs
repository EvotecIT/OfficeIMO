using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderCanvasTable(PdfCanvasTableItem item) {
            ValidateCanvasBox(item.X, item.Y, item.Width, item.Height, "Canvas table");
            TableBlock table = item.Block;
            int columns = GetTableColumnCount(table);
            int rows = table.Cells.Count;
            if (columns == 0 || rows == 0) {
                return;
            }

            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            ValidateCanvasTableStyle(style, rows, columns);

            double cellSpacing = GetTableCellSpacing(style);
            double columnGap = cellSpacing;
            double rowGap = cellSpacing;
            int headerRowCount = style.HeaderRowCount;
            int footerStart = rows - style.FooterRowCount;
            double[] columnWidths = ResolveCanvasTableColumnWidths(style, columns, item.Width, columnGap);
            double[] rowHeights = ResolveCanvasTableRowHeights(style, rows, item.Height, rowGap);
            var rowFontSizes = new double[rows];
            var rowFontSizeScales = new double[rows];
            var rowLeadings = new double[rows];
            for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
                bool rowUsesBold = GetTableRowBold(style, rowIndex, headerRowCount, footerStart);
                double originalRowFontSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStart, currentOpts.DefaultFontSize);
                double rowFontSize = ResolveTableRowShrinkFontSize(table, style, rowIndex, columns, columnWidths, columnGap, originalRowFontSize, rowUsesBold, currentOpts);
                rowFontSizes[rowIndex] = rowFontSize;
                rowFontSizeScales[rowIndex] = GetTableRunFontSizeScale(table, style, rowIndex, columns, columnWidths, columnGap, originalRowFontSize, rowFontSize, rowUsesBold, currentOpts);
                rowLeadings[rowIndex] = GetTableLeading(style, rowFontSize);
            }

            double tableWidth = GetTableCellWidth(columnWidths, 0, columns, columnGap);
            double tableHeight = GetTableRowsHeight(rowHeights, 0, rows, rowGap);
            double xOrigin = item.X;
            double topY = currentOpts.PageHeight - item.Y;
            double bottomY = topY - item.Height;
            int annotationStart = currentPage!.Annotations.Count;
            int imageStart = currentPage.Images.Count;
            int formFieldStart = currentPage.FormFields.Count;
            bool rotated = item.RotationAngle != 0D;
            if (rotated) {
                BeginRotatedCanvasFrame(item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            pageDirty = true;
            for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
                double rowTop = topY - GetTableRowsHeight(rowHeights, 0, rowIndex, rowGap);
                double rowBottom = rowTop - rowHeights[rowIndex];
                bool rowIsHeader = rowIndex < headerRowCount;
                bool rowIsFooter = rowIndex >= footerStart;
                bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(table, rowIndex, columns);
                DrawCanvasTableRowBackground(style, rowIndex, rowIsHeader, rowIsFooter, xOrigin, rowBottom, columnWidths, columnGap, rowFillSkips, rowHeights[rowIndex]);

                var cells = GetTableCellLayouts(table, rowIndex, columns);
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    double cellX = xOrigin + GetCanvasTableColumnsOffset(columnWidths, cell.Column, columnGap);
                    double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                    double cellHeight = GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGap);
                    double cellBottom = rowTop - cellHeight;

                    DrawCanvasTableCellBackground(style, rowIndex, cell.Column, rowIsHeader, rowIsFooter, cellX, cellBottom, cellWidth, cellHeight);
                }

                DrawTableCellDataBars(
                    sb,
                    style,
                    cells,
                    rowIndex,
                    columns,
                    xOrigin,
                    rowTop,
                    rowBottom,
                    rowHeights[rowIndex],
                    columnWidths,
                    columnGap,
                    rowHeights,
                    rowGap,
                    wholeRowSegment: true,
                    startLine: 0,
                    rowFillSkips,
                    artifact: true);
                DrawTableCellIcons(
                    sb,
                    style,
                    cells,
                    rowIndex,
                    columns,
                    xOrigin,
                    rowTop,
                    rowBottom,
                    rowHeights[rowIndex],
                    columnWidths,
                    columnGap,
                    rowHeights,
                    rowGap,
                    wholeRowSegment: true,
                    startLine: 0,
                    rowFillSkips,
                    artifact: true);
                for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                    TableCellLayout cell = cells[cellIndex];
                    double cellX = xOrigin + GetCanvasTableColumnsOffset(columnWidths, cell.Column, columnGap);
                    double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                    double cellHeight = GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGap);
                    double cellBottom = rowTop - cellHeight;
                    bool rowUsesBold = GetTableRowBold(style, rowIndex, headerRowCount, footerStart);

                    RenderCanvasTableCellText(item, style, cell, rowIndex, cell.Column, rowIsHeader, rowIsFooter, rowUsesBold, cellX, rowTop, cellBottom, cellWidth, cellHeight, rowFontSizes[rowIndex], rowLeadings[rowIndex], rowFontSizeScales[rowIndex], item.Y + GetTableRowsHeight(rowHeights, 0, rowIndex, rowGap));
                    DrawCanvasTableCellBorder(style, rowIndex, cell.Column, cellX, cellBottom, cellWidth, cellHeight);
                }
            }

            if (style.BorderColor is not null && style.BorderWidth > 0D) {
                DrawCanvasTableGrid(table, style, columns, rows, xOrigin, topY, tableHeight, columnWidths, rowHeights, columnGap, rowGap);
            }

            if (rotated) {
                new ContentStreamBuilder(sb)
                    .RestoreState();
                RotateCanvasPageImages(currentPage!.Images, imageStart, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
                RotateCanvasFormFields(currentPage.FormFields, formFieldStart, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
                RotateCanvasLinkAnnotations(currentPage!.Annotations, annotationStart, item.X, bottomY, item.Width, item.Height, item.RotationAngle);
            }

            DrawDebugCanvasItemBox(item.X, bottomY, item.Width, item.Height);
        }

        private static void ValidateCanvasTableStyle(PdfTableStyle style, int rows, int columns) {
            ValidateTableRoleRowCounts(style, rows);
            ValidateTableCellStyleCoordinates(style, rows, columns);
            ValidateTableColumnStyleBounds(style, columns);
            ValidateTableRowStyleBounds(style, rows);
            if (style.CellSpacing * Math.Max(0, columns - 1) >= double.MaxValue ||
                style.CellSpacing * Math.Max(0, rows - 1) >= double.MaxValue) {
                throw new ArgumentException("Canvas table cell spacing must be finite.");
            }
        }

        private static double[] ResolveCanvasTableColumnWidths(PdfTableStyle style, int columns, double width, double columnGap) {
            double innerWidth = width - Math.Max(0, columns - 1) * columnGap;
            if (innerWidth <= 0.001D || double.IsNaN(innerWidth) || double.IsInfinity(innerWidth)) {
                throw new ArgumentException("Canvas table cell spacing must leave a positive table width.");
            }

            var widths = new double[columns];
            double total = 0D;
            bool allFixed = style.ColumnWidthPoints != null && style.ColumnWidthPoints.Count >= columns;
            for (int column = 0; column < columns; column++) {
                double? fixedWidth = style.ColumnWidthPoints != null &&
                    column < style.ColumnWidthPoints.Count &&
                    style.ColumnWidthPoints[column].HasValue
                        ? style.ColumnWidthPoints[column]!.Value
                        : (double?)null;
                if (!fixedWidth.HasValue) {
                    allFixed = false;
                    break;
                }

                widths[column] = fixedWidth.Value;
                total += fixedWidth.Value;
            }

            if (allFixed && total > 0D) {
                double scale = innerWidth / total;
                for (int column = 0; column < columns; column++) {
                    widths[column] *= scale;
                }

                return widths;
            }

            double equalWidth = innerWidth / columns;
            for (int column = 0; column < columns; column++) {
                widths[column] = equalWidth;
            }

            return widths;
        }

        private static double[] ResolveCanvasTableRowHeights(PdfTableStyle style, int rows, double height, double rowGap) {
            double innerHeight = height - Math.Max(0, rows - 1) * rowGap;
            if (innerHeight <= 0.001D || double.IsNaN(innerHeight) || double.IsInfinity(innerHeight)) {
                throw new ArgumentException("Canvas table cell spacing must leave a positive table height.");
            }

            var heights = new double[rows];
            double total = 0D;
            bool allFixed = (style.FixedRowHeights != null && style.FixedRowHeights.Count >= rows) ||
                (style.RowMinHeights != null && style.RowMinHeights.Count >= rows);
            for (int row = 0; row < rows; row++) {
                double? fixedHeight = GetTableRowFixedHeight(style, row);
                if (!fixedHeight.HasValue &&
                    style.RowMinHeights != null &&
                    row < style.RowMinHeights.Count &&
                    style.RowMinHeights[row].HasValue &&
                    style.RowMinHeights[row]!.Value > 0D) {
                    fixedHeight = style.RowMinHeights[row]!.Value;
                }

                if (!fixedHeight.HasValue) {
                    allFixed = false;
                    break;
                }

                heights[row] = fixedHeight.Value;
                total += fixedHeight.Value;
            }

            if (allFixed && total > 0D) {
                double scale = innerHeight / total;
                for (int row = 0; row < rows; row++) {
                    heights[row] *= scale;
                }

                return heights;
            }

            double equalHeight = innerHeight / rows;
            for (int row = 0; row < rows; row++) {
                heights[row] = equalHeight;
            }

            return heights;
        }

        private void DrawCanvasTableRowBackground(PdfTableStyle style, int rowIndex, bool rowIsHeader, bool rowIsFooter, double x, double y, double[] columnWidths, double columnGap, bool[] skipColumns, double height) {
            PdfColor? fill = rowIsHeader
                ? style.HeaderFill
                : rowIsFooter
                    ? style.FooterFill
                    : ((rowIndex - style.HeaderRowCount) % 2 == 1 ? style.RowStripeFill : null);
            if (fill.HasValue) {
                DrawTableRowFill(sb, fill.Value, x, columnWidths, columnGap, y, height, skipColumns, true);
            }
        }

        private void DrawCanvasTableCellBackground(PdfTableStyle style, int rowIndex, int columnIndex, bool rowIsHeader, bool rowIsFooter, double x, double y, double width, double height) {
            PdfColor? fill = null;
            if (style.BodyColumnFills != null &&
                !rowIsHeader &&
                !rowIsFooter &&
                columnIndex < style.BodyColumnFills.Count) {
                fill = style.BodyColumnFills[columnIndex];
            }

            if (style.CellFills != null && style.CellFills.TryGetValue((rowIndex, columnIndex), out PdfColor cellFill)) {
                fill = cellFill;
            }

            if (fill.HasValue) {
                DrawRowFill(sb, fill.Value, x, y, width, height, true);
            }
        }

        private void RenderCanvasTableCellText(PdfCanvasTableItem item, PdfTableStyle style, TableCellLayout cell, int rowIndex, int columnIndex, bool rowIsHeader, bool rowIsFooter, bool rowUsesBold, double cellX, double cellTop, double cellBottom, double cellWidth, double cellHeight, double fontSize, double leading, double runFontSizeScale, double cellYFromTop) {
            PdfStandardFont cellFont = GetTableRowFont(currentOpts, rowUsesBold);
            double padLeft = GetTableCellPaddingLeft(style, rowIndex, columnIndex);
            double padRight = GetTableCellPaddingRight(style, rowIndex, columnIndex);
            double padTop = GetTableCellPaddingTop(style, rowIndex, columnIndex);
            double padBottom = GetTableCellPaddingBottom(style, rowIndex, columnIndex);
            double innerWidth = Math.Max(1D, cellWidth - padLeft - padRight);
            double availableHeight = Math.Max(0D, cellHeight - padTop - padBottom);
            var lines = CreateTableCellTextLayout(cell, innerWidth, cellFont, fontSize, leading, currentOpts, runFontSizeScale, style.MinimumShrinkFontSize ?? 6D);
            int lineCount = Math.Max(1, lines.LineCount);
            double contentHeight = MeasureTableCellContentHeight(cell, lines, 0, lineCount, leading, innerWidth);
            if (contentHeight > availableHeight + 0.01D) {
                item.DiagnosticHandler?.Invoke(new PdfLayoutDiagnostic(
                    PdfLayoutDiagnosticKind.ClippedContent,
                    "PdfCanvasTableCell",
                    "The PDF table render pass clipped cell text because wrapped content exceeded the available cell area.",
                    cellX,
                    cellYFromTop,
                    cellWidth,
                    cellHeight));
            }

            double verticalOffset = 0D;
            PdfCellVerticalAlign verticalAlign = GetTableCellVerticalAlignment(style, rowIndex, columnIndex);
            if (availableHeight > contentHeight) {
                if (verticalAlign == PdfCellVerticalAlign.Middle) {
                    verticalOffset = (availableHeight - contentHeight) / 2D;
                } else if (verticalAlign == PdfCellVerticalAlign.Bottom) {
                    verticalOffset = availableHeight - contentHeight;
                }
            }

            double firstBaseline = cellTop - padTop - verticalOffset - GetAscenderForOptions(cellFont, fontSize, currentOpts) + style.RowBaselineOffset;
            var visibleLines = SliceTableCellLines(lines, 0, lineCount);
            var visibleHeights = SliceTableCellLineHeights(lines, 0, lineCount, leading);
            var visibleAlignments = SliceTableCellLineAlignments(lines, 0, lineCount);
            var visibleXOffsets = SliceTableCellLineXOffsets(lines, 0, lineCount);
            var visibleWidths = SliceTableCellLineWidths(lines, 0, lineCount, innerWidth);
            string? linkUri = cell.LinkUri;
            string? linkDestinationName = cell.LinkDestinationName;
            string? linkContents = cell.LinkContents;
            if (cell.LinkUri != null || cell.LinkDestinationName != null) {
                visibleLines = StripRichLineLinksWhenCellLinked(visibleLines, linkUri, linkDestinationName);
            }

            PdfColumnAlign align = GetTableCellAlignment(style, rowIndex, columnIndex, cell.Text);
            PdfColor? textColor = rowIsHeader ? style.HeaderTextColor : rowIsFooter ? style.FooterTextColor : style.TextColor;
            var paragraph = new RichParagraphBlock(StripRunLinksWhenCellLinked(cell.Runs, linkUri, linkDestinationName), MapTableCellAlignment(align), textColor);
            int? markedContentId = RegisterTextStructureElement(rowIsHeader ? "TH" : "TD", _canvasStructureParentElementIndex);
            WriteClippedRichParagraph(
                sb,
                paragraph,
                visibleLines,
                visibleHeights,
                currentOpts,
                firstBaseline,
                fontSize,
                leading,
                currentPage!.Annotations,
                cellX - TableCellClipBleed,
                cellBottom - TableCellClipBleed,
                cellWidth + (TableCellClipBleed * 2D),
                cellHeight + (TableCellClipBleed * 2D),
                cellX + padLeft,
                innerWidth,
                structureType: rowIsHeader ? "TH" : "TD",
                markedContentId: markedContentId,
                structurePage: currentPage,
                lineAlignments: visibleAlignments,
                lineXOffsets: visibleXOffsets,
                lineWidths: visibleWidths);
            if (cell.Runs.Any(run => run.Bold || rowUsesBold)) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            if (cell.Runs.Any(run => run.Italic)) {
                currentPage!.UsedItalic = true;
                usedItalic = true;
            }

            if (cell.Runs.Any(run => (run.Bold || rowUsesBold) && run.Italic)) {
                currentPage!.UsedBoldItalic = true;
                usedBoldItalic = true;
            }

            MarkRichFonts(cell.Runs);
            AddTableCellNamedDestinationName(cell.NamedDestinationName, cellTop);
            if (cell.Images.Count > 0 || cell.CheckBoxes.Count > 0 || cell.FormFields.Count > 0) {
                if (CanRenderTableCellCheckBoxInline(cell, lines, 0, lineCount)) {
                    RenderTableCellInlineCheckBox(currentPage!, cell, align, lines.Lines[0], cellX + padLeft, innerWidth, firstBaseline);
                } else {
                    double textHeight = MeasureTableCellTextHeight(lines, 0, lineCount, leading);
                    double formFieldTop = cellTop - padTop - verticalOffset - (string.IsNullOrEmpty(cell.Text) ? 0D : textHeight + TableCellCheckBoxGap);
                    RenderTableCellObjects(currentPage!, cell, align, cellX + padLeft, innerWidth, formFieldTop);
                }
            }

            if (HasCellLinkTarget(linkUri, linkDestinationName)) {
                currentPage!.Annotations.Add(new LinkAnnotation {
                    X1 = cellX + padLeft - TableCellClipBleed,
                    Y1 = cellBottom - TableCellClipBleed,
                    X2 = cellX + cellWidth - padRight + TableCellClipBleed,
                    Y2 = cellTop + TableCellClipBleed,
                    Uri = linkUri,
                    DestinationName = linkDestinationName,
                    Contents = linkContents ?? cell.Text
                });
            }
        }

        private void DrawCanvasTableCellBorder(PdfTableStyle style, int rowIndex, int columnIndex, double x, double y, double width, double height) {
            if (style.CellBorders != null &&
                style.CellBorders.TryGetValue((rowIndex, columnIndex), out PdfCellBorder? border) &&
                HasRenderableCellBorder(border)) {
                DrawCellBorder(sb, border, x, y, width, height, true);
            }
        }

        private void DrawCanvasTableGrid(TableBlock table, PdfTableStyle style, int columns, int rows, double x, double topY, double height, double[] columnWidths, double[] rowHeights, double columnGap, double rowGap) {
            PdfColor color = style.BorderColor!.Value;
            double width = style.BorderWidth;
            double tableWidth = GetTableCellWidth(columnWidths, 0, columns, columnGap);
            DrawRowRect(sb, color, width, x, topY - height, tableWidth, height, true);

            double lineX = x;
            for (int column = 0; column < columns - 1; column++) {
                lineX += columnWidths[column];
                for (int row = 0; row < rows; row++) {
                    if (IsTableBoundaryInsideSpannedCell(table, row, column, columns)) {
                        continue;
                    }

                    double lineTop = topY - GetTableRowsHeight(rowHeights, 0, row, rowGap);
                    double lineBottom = lineTop - rowHeights[row];
                    DrawVLine(sb, color, width, lineX, lineTop, lineBottom, true);
                }

                lineX += columnGap;
            }

            double lineY = topY;
            for (int row = 0; row < rows - 1; row++) {
                lineY -= rowHeights[row];
                bool[] skips = GetRowSpanBoundarySkipColumns(table, row, columns);
                DrawTableHorizontalLine(sb, color, width, x, columnWidths, columnGap, lineY, skips, true);
                lineY -= rowGap;
            }
        }

        private static double GetCanvasTableColumnsOffset(double[] columnWidths, int column, double columnGap) {
            double offset = 0D;
            for (int index = 0; index < column; index++) {
                offset += columnWidths[index] + columnGap;
            }

            return offset;
        }
    }
}
