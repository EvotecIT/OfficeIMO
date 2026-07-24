using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double GetTableBodyFontSize(PdfTableStyle style, double defaultFontSize) {
        return style.FontSize ?? defaultFontSize;
    }

    private static double GetTableLeading(PdfTableStyle style, double fontSize) {
        double multiplier = style.LineHeight ?? 1.4D;
        if (multiplier <= 0 || double.IsNaN(multiplier) || double.IsInfinity(multiplier)) {
            throw new ArgumentException("Table line height must be a positive finite value.");
        }

        return fontSize * multiplier;
    }

    private static double GetTableCellPaddingLeft(PdfTableStyle style) {
        return style.CellPaddingLeft ?? style.CellPaddingX;
    }

    private static double GetTableCellPaddingRight(PdfTableStyle style) {
        return style.CellPaddingRight ?? style.CellPaddingX;
    }

    private static double GetTableCellPaddingTop(PdfTableStyle style) {
        return style.CellPaddingTop ?? style.CellPaddingY;
    }

    private static double GetTableCellPaddingBottom(PdfTableStyle style) {
        return style.CellPaddingBottom ?? style.CellPaddingY;
    }

    private static PdfCellPadding? GetTableCellPaddingOverride(PdfTableStyle style, int rowIndex, int columnIndex) {
        if (style.CellPaddings != null &&
            style.CellPaddings.TryGetValue((rowIndex, columnIndex), out PdfCellPadding? padding)) {
            return padding;
        }

        return null;
    }

    private static double GetTableCellPaddingLeft(PdfTableStyle style, int rowIndex, int columnIndex) {
        return GetTableCellPaddingOverride(style, rowIndex, columnIndex)?.Left ?? GetTableCellPaddingLeft(style);
    }

    private static double GetTableCellPaddingRight(PdfTableStyle style, int rowIndex, int columnIndex) {
        return GetTableCellPaddingOverride(style, rowIndex, columnIndex)?.Right ?? GetTableCellPaddingRight(style);
    }

    private static double GetTableCellPaddingTop(PdfTableStyle style, int rowIndex, int columnIndex) {
        return GetTableCellPaddingOverride(style, rowIndex, columnIndex)?.Top ?? GetTableCellPaddingTop(style);
    }

    private static double GetTableCellPaddingBottom(PdfTableStyle style, int rowIndex, int columnIndex) {
        return GetTableCellPaddingOverride(style, rowIndex, columnIndex)?.Bottom ?? GetTableCellPaddingBottom(style);
    }

    private static double GetTableRowMaxPaddingTop(TableBlock table, PdfTableStyle style, int rowIndex, int columnCount) {
        double padding = GetTableCellPaddingTop(style);
        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            padding = Math.Max(padding, GetTableCellPaddingTop(style, rowIndex, cell.Column));
        }

        return padding;
    }

    private static double GetTableRowMaxPaddingBottom(TableBlock table, PdfTableStyle style, int rowIndex, int columnCount) {
        double padding = GetTableCellPaddingBottom(style);
        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            padding = Math.Max(padding, GetTableCellPaddingBottom(style, rowIndex, cell.Column));
        }

        return padding;
    }

    private static double GetTableCellSpacing(PdfTableStyle style) {
        double spacing = style.CellSpacing;
        if (spacing < 0 || double.IsNaN(spacing) || double.IsInfinity(spacing)) {
            throw new ArgumentException("Table cell spacing must be a non-negative finite value.");
        }

        return spacing;
    }

    private static PdfColumnAlign GetTableCellAlignment(PdfTableStyle style, int rowIndex, int columnIndex, string cellText) {
        if (style.CellAlignments != null &&
            style.CellAlignments.TryGetValue((rowIndex, columnIndex), out PdfColumnAlign cellAlignment)) {
            return cellAlignment;
        }

        var alignment = PdfColumnAlign.Left;
        if (style.Alignments != null && columnIndex < style.Alignments.Count) {
            alignment = style.Alignments[columnIndex];
        }

        if (style.RightAlignNumeric && LooksNumeric(cellText)) {
            return PdfColumnAlign.Right;
        }

        return alignment;
    }

    private static PdfCellVerticalAlign GetTableCellVerticalAlignment(PdfTableStyle style, int rowIndex, int columnIndex) {
        if (style.CellVerticalAlignments != null &&
            style.CellVerticalAlignments.TryGetValue((rowIndex, columnIndex), out PdfCellVerticalAlign cellAlignment)) {
            return cellAlignment;
        }

        if (style.VerticalAlignments != null && columnIndex < style.VerticalAlignments.Count) {
            return style.VerticalAlignments[columnIndex];
        }

        return PdfCellVerticalAlign.Top;
    }

    private static double GetTableRowMinHeight(PdfTableStyle style, int rowIndex) {
        if (style.RowMinHeights != null &&
            rowIndex < style.RowMinHeights.Count &&
            style.RowMinHeights[rowIndex].HasValue) {
            return style.RowMinHeights[rowIndex]!.Value;
        }

        return style.MinRowHeight;
    }

    private static double? GetTableRowFixedHeight(PdfTableStyle style, int rowIndex) {
        if (style.FixedRowHeights != null &&
            rowIndex < style.FixedRowHeights.Count &&
            style.FixedRowHeights[rowIndex].HasValue) {
            return style.FixedRowHeights[rowIndex]!.Value;
        }

        return null;
    }

    private static double ResolveTableRowShrinkFontSize(TableBlock table, PdfTableStyle style, int rowIndex, int columnCount, double[] columnWidths, double columnGap, double rowFontSize, bool rowUsesBold, PdfOptions? options) {
        if (!style.ShrinkTextToFit || rowFontSize <= 0D) {
            return rowFontSize;
        }

        double minimumFontSize = style.MinimumShrinkFontSize ?? 6D;
        if (minimumFontSize > rowFontSize) {
            return rowFontSize;
        }

        PdfStandardFont rowFont = GetTableRowFont(options ?? new PdfOptions(), rowUsesBold);
        double resolvedFontSize = rowFontSize;
        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            if (string.IsNullOrEmpty(cell.Text)) {
                continue;
            }

            double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
            double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
            double textWidth = MeasureTableCellTextWidth(cell, rowFont, rowFontSize, options);
            if (textWidth <= innerWidth + 0.001D || textWidth <= 0.001D) {
                continue;
            }

            double candidate = Math.Max(minimumFontSize, rowFontSize * innerWidth / textWidth);
            resolvedFontSize = Math.Min(resolvedFontSize, candidate);
        }

        return resolvedFontSize;
    }

    private static double MeasureTableCellTextWidth(TableCellLayout cell, PdfStandardFont baseFont, double fontSize, PdfOptions? options) {
        double width = 0D;
        if (cell.Paragraphs.Count > 0) {
            foreach (PdfTableCellParagraph paragraph in cell.Paragraphs) {
                width = Math.Max(width, MeasureTableRunsTextWidth(paragraph.Runs, baseFont, fontSize, options));
            }
        } else {
            width = MeasureTableRunsTextWidth(cell.Runs, baseFont, fontSize, options);
        }

        return width;
    }

    private static double MeasureTableRunsTextWidth(System.Collections.Generic.IReadOnlyList<TextRun> runs, PdfStandardFont baseFont, double fontSize, PdfOptions? options) {
        PdfOptions effectiveOptions = options ?? new PdfOptions();
        System.Collections.Generic.IReadOnlyList<TextRun> normalizedRuns = NormalizeFallbackRuns(runs, baseFont, effectiveOptions);
        double width = 0D;
        foreach (System.Collections.Generic.IReadOnlyList<TextRun> line in BuildPageTextLineRuns(normalizedRuns)) {
            width = Math.Max(width, MeasurePageTextLineRuns(line, baseFont, fontSize, effectiveOptions));
        }

        return width;
    }

    private static double GetTableRunFontSizeScale(double originalFontSize, double resolvedFontSize) {
        if (originalFontSize <= 0D ||
            resolvedFontSize >= originalFontSize - 0.001D) {
            return 1D;
        }

        return resolvedFontSize / originalFontSize;
    }

    private static double GetTableRunFontSizeScale(TableBlock table, PdfTableStyle style, int rowIndex, int columnCount, double[] columnWidths, double columnGap, double originalFontSize, double resolvedFontSize, bool rowUsesBold, PdfOptions? options) {
        double scale = GetTableRunFontSizeScale(originalFontSize, resolvedFontSize);
        if (!style.ShrinkTextToFit) {
            return scale;
        }

        double minimumFontSize = style.MinimumShrinkFontSize ?? 6D;
        PdfStandardFont rowFont = GetTableRowFont(options ?? new PdfOptions(), rowUsesBold);
        var cells = GetTableCellLayouts(table, rowIndex, columnCount);
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            TableCellLayout cell = cells[cellIndex];
            double maxExplicitFontSize = GetMaxExplicitTableRunFontSize(cell);
            if (maxExplicitFontSize <= resolvedFontSize + 0.001D) {
                continue;
            }

            double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
            double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
            double textWidth = MeasureTableCellTextWidth(cell, rowFont, resolvedFontSize, options, scale, minimumFontSize);
            if (textWidth <= innerWidth + 0.001D || textWidth <= 0.001D) {
                continue;
            }

            double minimumScale = 0.001D;
            double minimumWidth = MeasureTableCellTextWidth(cell, rowFont, resolvedFontSize, options, minimumScale, minimumFontSize);
            if (minimumWidth > innerWidth + 0.001D) {
                scale = Math.Min(scale, minimumScale);
                continue;
            }

            double low = minimumScale;
            double high = scale;
            for (int iteration = 0; iteration < 20; iteration++) {
                double candidate = (low + high) / 2D;
                double candidateWidth = MeasureTableCellTextWidth(cell, rowFont, resolvedFontSize, options, candidate, minimumFontSize);
                if (candidateWidth <= innerWidth + 0.001D) {
                    low = candidate;
                } else {
                    high = candidate;
                }
            }

            scale = Math.Min(scale, low);
        }

        return scale;
    }

    private static double MeasureTableCellTextWidth(TableCellLayout cell, PdfStandardFont baseFont, double fontSize, PdfOptions? options, double runFontSizeScale, double minimumShrinkFontSize) {
        if (runFontSizeScale >= 0.999D) {
            return MeasureTableCellTextWidth(cell, baseFont, fontSize, options);
        }

        double width = 0D;
        if (cell.Paragraphs.Count > 0) {
            foreach (PdfTableCellParagraph paragraph in cell.Paragraphs) {
                width = Math.Max(width, MeasureTableRunsTextWidth(ScaleTableRunsForShrink(paragraph.Runs, runFontSizeScale, minimumShrinkFontSize), baseFont, fontSize, options));
            }
        } else {
            width = MeasureTableRunsTextWidth(ScaleTableRunsForShrink(cell.Runs, runFontSizeScale, minimumShrinkFontSize), baseFont, fontSize, options);
        }

        return width;
    }

    private static double GetMaxExplicitTableRunFontSize(TableCellLayout cell) {
        double max = GetMaxExplicitRunFontSize(cell.Runs);
        for (int i = 0; i < cell.Paragraphs.Count; i++) {
            max = Math.Max(max, GetMaxExplicitRunFontSize(cell.Paragraphs[i].Runs));
        }

        return max;
    }

    private static double GetMaxExplicitRunFontSize(System.Collections.Generic.IReadOnlyList<TextRun> runs) {
        double max = 0D;
        foreach (TextRun run in runs) {
            if (run.FontSize.HasValue) {
                max = Math.Max(max, run.FontSize.Value);
            }
        }

        return max;
    }

    private static double ResolveTableRowHeight(PdfTableStyle style, int rowIndex, double requiredHeight) {
        double? fixedHeight = GetTableRowFixedHeight(style, rowIndex);
        return fixedHeight ?? System.Math.Max(requiredHeight, GetTableRowMinHeight(style, rowIndex));
    }

    private static bool GetTableRowAllowBreakAcrossPages(PdfTableStyle style, int rowIndex) {
        if (style.RowAllowBreakAcrossPages != null &&
            rowIndex < style.RowAllowBreakAcrossPages.Count &&
            style.RowAllowBreakAcrossPages[rowIndex].HasValue) {
            return style.RowAllowBreakAcrossPages[rowIndex]!.Value;
        }

        return style.AllowRowBreakAcrossPages;
    }

    private static int GetTableColumnCount(TableBlock table) => table.ColumnCount;

    private static void ValidateTableRoleRowCounts(PdfTableStyle style, int rowCount) {
        if (style.HeaderRowCount > rowCount) {
            throw new ArgumentException("Table header row count cannot exceed the table row count.");
        }

        int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
        if (repeatHeaderRowCount > style.HeaderRowCount) {
            throw new ArgumentException("Table repeating header row count cannot exceed the table header row count.");
        }

        if (style.FooterRowCount > rowCount) {
            throw new ArgumentException("Table footer row count cannot exceed the table row count.");
        }

        if (style.FooterRowCount > rowCount - style.HeaderRowCount) {
            throw new ArgumentException("Table header and footer row counts cannot exceed the table row count.");
        }
    }

    private static void ValidateTableCellStyleCoordinates(PdfTableStyle style, TableBlock table, int columnCount) {
        int rowCount = table.Rows.Count;
        if (style.CellFills != null) {
            foreach (var cellFill in style.CellFills) {
                if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                    throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellFill.Key.Row, cellFill.Key.Column, "Table cell fill coordinates must fit inside the table grid.");
            }
        }

        if (style.CellBorders != null) {
            foreach (var cellBorder in style.CellBorders) {
                if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                    throw new ArgumentException("Table cell border coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellBorder.Key.Row, cellBorder.Key.Column, "Table cell border coordinates must fit inside the table grid.");
            }
        }

        if (style.CellDataBars != null) {
            foreach (var cellDataBar in style.CellDataBars) {
                if (cellDataBar.Key.Row < 0 || cellDataBar.Key.Column < 0) {
                    throw new ArgumentException("Table cell data bar coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellDataBar.Key.Row, cellDataBar.Key.Column, "Table cell data bar coordinates must fit inside the table grid.");
            }
        }

        if (style.CellIcons != null) {
            foreach (var cellIcon in style.CellIcons) {
                if (cellIcon.Key.Row < 0 || cellIcon.Key.Column < 0) {
                    throw new ArgumentException("Table cell icon coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellIcon.Key.Row, cellIcon.Key.Column, "Table cell icon coordinates must fit inside the table grid.");
            }
        }

        if (style.CellPaddings != null) {
            foreach (var cellPadding in style.CellPaddings) {
                if (cellPadding.Key.Row < 0 || cellPadding.Key.Column < 0) {
                    throw new ArgumentException("Table cell padding coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellPadding.Key.Row, cellPadding.Key.Column, "Table cell padding coordinates must fit inside the table grid.");
            }
        }

        if (style.CellAlignments != null) {
            foreach (var cellAlignment in style.CellAlignments) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new ArgumentException("Table cell alignment coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellAlignment.Key.Row, cellAlignment.Key.Column, "Table cell alignment coordinates must fit inside the table grid.");

                if (!IsValidPdfColumnAlign(cellAlignment.Value)) {
                    throw new ArgumentException("Table cell alignments must be Left, Center, or Right.");
                }
            }
        }

        if (style.CellVerticalAlignments != null) {
            foreach (var cellAlignment in style.CellVerticalAlignments) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new ArgumentException("Table cell vertical alignment coordinates cannot be negative.");
                }

                ValidateTableCellStyleAnchor(table, columnCount, cellAlignment.Key.Row, cellAlignment.Key.Column, "Table cell vertical alignment coordinates must fit inside the table grid.");

                if (!IsValidPdfCellVerticalAlign(cellAlignment.Value)) {
                    throw new ArgumentException("Table cell vertical alignments must be defined PDF cell vertical alignment values.");
                }
            }
        }
    }

    private static void ValidateTableCellStyleCoordinates(PdfTableStyle style, int rowCount, int columnCount) {
        if (style.CellFills != null) {
            foreach (var cellFill in style.CellFills) {
                if (cellFill.Key.Row < 0 || cellFill.Key.Column < 0) {
                    throw new ArgumentException("Table cell fill coordinates cannot be negative.");
                }

                if (cellFill.Key.Row >= rowCount || cellFill.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell fill coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellBorders != null) {
            foreach (var cellBorder in style.CellBorders) {
                if (cellBorder.Key.Row < 0 || cellBorder.Key.Column < 0) {
                    throw new ArgumentException("Table cell border coordinates cannot be negative.");
                }

                if (cellBorder.Key.Row >= rowCount || cellBorder.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell border coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellDataBars != null) {
            foreach (var cellDataBar in style.CellDataBars) {
                if (cellDataBar.Key.Row < 0 || cellDataBar.Key.Column < 0) {
                    throw new ArgumentException("Table cell data bar coordinates cannot be negative.");
                }

                if (cellDataBar.Key.Row >= rowCount || cellDataBar.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell data bar coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellIcons != null) {
            foreach (var cellIcon in style.CellIcons) {
                if (cellIcon.Key.Row < 0 || cellIcon.Key.Column < 0) {
                    throw new ArgumentException("Table cell icon coordinates cannot be negative.");
                }

                if (cellIcon.Key.Row >= rowCount || cellIcon.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell icon coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellPaddings != null) {
            foreach (var cellPadding in style.CellPaddings) {
                if (cellPadding.Key.Row < 0 || cellPadding.Key.Column < 0) {
                    throw new ArgumentException("Table cell padding coordinates cannot be negative.");
                }

                if (cellPadding.Key.Row >= rowCount || cellPadding.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell padding coordinates must fit inside the table grid.");
                }
            }
        }

        if (style.CellAlignments != null) {
            foreach (var cellAlignment in style.CellAlignments) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new ArgumentException("Table cell alignment coordinates cannot be negative.");
                }

                if (cellAlignment.Key.Row >= rowCount || cellAlignment.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell alignment coordinates must fit inside the table grid.");
                }

                if (!IsValidPdfColumnAlign(cellAlignment.Value)) {
                    throw new ArgumentException("Table cell alignments must be Left, Center, or Right.");
                }
            }
        }

        if (style.CellVerticalAlignments != null) {
            foreach (var cellAlignment in style.CellVerticalAlignments) {
                if (cellAlignment.Key.Row < 0 || cellAlignment.Key.Column < 0) {
                    throw new ArgumentException("Table cell vertical alignment coordinates cannot be negative.");
                }

                if (cellAlignment.Key.Row >= rowCount || cellAlignment.Key.Column >= columnCount) {
                    throw new ArgumentException("Table cell vertical alignment coordinates must fit inside the table grid.");
                }

                if (!IsValidPdfCellVerticalAlign(cellAlignment.Value)) {
                    throw new ArgumentException("Table cell vertical alignments must be defined PDF cell vertical alignment values.");
                }
            }
        }
    }

    private static void ValidateTableCellStyleAnchor(TableBlock table, int columnCount, int row, int column, string outOfRangeMessage) {
        if (row >= table.Rows.Count || column >= columnCount) {
            throw new ArgumentException(outOfRangeMessage);
        }
    }

    private static int GetTableRepeatHeaderRowCount(PdfTableStyle style) =>
        style.RepeatHeaderRowCount ?? style.HeaderRowCount;

    private static bool ShouldBreakBeforeFinalTableBodyRows(
        int rowIndex,
        int headerRowCount,
        int footerStartRowIndex,
        int minimumBodyRows,
        double firstRowHeight,
        double finalGroupHeight,
        double remainingHeight,
        double continuationHeaderHeight,
        double maxContentHeight,
        bool hasContentBeforeRow) {
        if (!hasContentBeforeRow ||
            minimumBodyRows <= 0 ||
            rowIndex < headerRowCount ||
            rowIndex >= footerStartRowIndex ||
            footerStartRowIndex - rowIndex != minimumBodyRows) {
            return false;
        }

        if (firstRowHeight > remainingHeight + 0.001 ||
            finalGroupHeight <= remainingHeight + 0.001) {
            return false;
        }

        return continuationHeaderHeight + finalGroupHeight <= maxContentHeight + 0.001;
    }

    private static void ValidateTableColumnStyleBounds(PdfTableStyle style, int columnCount) {
        if (style.BodyColumnFills != null) {
            for (int column = columnCount; column < style.BodyColumnFills.Count; column++) {
                if (style.BodyColumnFills[column] != null) {
                    throw new ArgumentException("Table body column fills must fit inside the table grid.");
                }
            }
        }

        if (style.Alignments != null && style.Alignments.Count > columnCount) {
            throw new ArgumentException("Table column alignments must fit inside the table grid.");
        }

        if (style.VerticalAlignments != null && style.VerticalAlignments.Count > columnCount) {
            throw new ArgumentException("Table vertical alignments must fit inside the table grid.");
        }

        ValidateOptionalColumnStyleBounds(style.ColumnWidthPoints, columnCount, "Table fixed column widths must fit inside the table grid.");
        ValidateOptionalColumnStyleBounds(style.ColumnMinWidthPoints, columnCount, "Table minimum column widths must fit inside the table grid.");
        ValidateOptionalColumnStyleBounds(style.ColumnMaxWidthPoints, columnCount, "Table maximum column widths must fit inside the table grid.");

        if (style.ColumnWidthWeights != null && style.ColumnWidthWeights.Count > columnCount) {
            throw new ArgumentException("Table column width weights must fit inside the table grid.");
        }
    }

    private static void ValidateOptionalColumnStyleBounds(System.Collections.Generic.List<double?>? values, int columnCount, string message) {
        if (values == null) {
            return;
        }

        for (int column = columnCount; column < values.Count; column++) {
            if (values[column].HasValue) {
                throw new ArgumentException(message);
            }
        }
    }

    private static System.Collections.Generic.List<TableCellLayout> GetTableCellLayouts(TableBlock table, int rowIndex, int columnCount) {
        var targetCells = new System.Collections.Generic.List<TableCellLayout>();
        if (rowIndex < 0 || rowIndex >= table.Cells.Count) {
            return targetCells;
        }

        var activeRowSpans = new int[columnCount];
        for (int currentRow = 0; currentRow <= rowIndex; currentRow++) {
            int column = 0;
            var row = table.Cells[currentRow];
            for (int cellIndex = 0; cellIndex < row.Count && column < columnCount; cellIndex++) {
                while (column < columnCount && activeRowSpans[column] > 0) {
                    column++;
                }

                if (column >= columnCount) {
                    break;
                }

                PdfTableCell cell = row[cellIndex];
                int columnSpan = System.Math.Min(cell.ColumnSpan, columnCount - column);
                int rowSpan = System.Math.Min(cell.RowSpan, table.Cells.Count - currentRow);
                if (currentRow == rowIndex) {
                    targetCells.Add(new TableCellLayout(column, columnSpan, rowSpan, cell.Text, cell.Runs, cell.Paragraphs, cell.LinkUri, cell.LinkDestinationName, cell.LinkContents, cell.NamedDestinationName, cell.CheckBoxes, cell.FormFields, cell.Images, cell.NoWrap));
                }

                for (int c = column; c < column + columnSpan; c++) {
                    activeRowSpans[c] = System.Math.Max(activeRowSpans[c], rowSpan);
                }

                column += columnSpan;
            }

            for (int c = 0; c < activeRowSpans.Length; c++) {
                if (activeRowSpans[c] > 0) {
                    activeRowSpans[c]--;
                }
            }
        }

        return targetCells;
    }

    private static double GetTableCellWidth(double[] columnWidths, int column, int columnSpan, double columnGap) {
        double width = 0D;
        int lastColumn = System.Math.Min(columnWidths.Length, column + columnSpan);
        for (int index = column; index < lastColumn; index++) {
            width += columnWidths[index];
            if (index > column) {
                width += columnGap;
            }
        }

        return width;
    }

    private static double GetTableCellHeight(double[] rowHeights, int row, int rowSpan, double rowGap = 0D) {
        double height = 0D;
        int lastRow = System.Math.Min(rowHeights.Length, row + rowSpan);
        for (int index = row; index < lastRow; index++) {
            height += rowHeights[index];
            if (index > row) {
                height += rowGap;
            }
        }

        return height;
    }

    private static double GetTableRowGapAfter(int rowIndex, int rowCount, double rowGap) =>
        rowIndex < rowCount - 1 ? rowGap : 0D;

    private static double GetTableRowsHeight(double[] rowHeights, int startRow, int rowCount, double rowGap) {
        double height = 0D;
        int lastRow = System.Math.Min(rowHeights.Length, startRow + rowCount);
        for (int rowIndex = startRow; rowIndex < lastRow; rowIndex++) {
            height += rowHeights[rowIndex] + GetTableRowGapAfter(rowIndex, rowHeights.Length, rowGap);
        }

        return height;
    }

    private static TableCellTextLayout CreateTableCellTextLayout(TableCellLayout cell, double innerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options, double runFontSizeScale = 1D, double minimumShrinkFontSize = 0D) {
        double wrapWidth = GetTableCellWrapWidth(innerWidth, cell.NoWrap);
        if (cell.Paragraphs.Count > 0) {
            return CreateTableCellParagraphTextLayout(ScaleTableCellParagraphsForShrink(cell.Paragraphs, runFontSizeScale, minimumShrinkFontSize), wrapWidth, innerWidth, baseFont, fontSize, leading, options);
        }

        var wrap = WrapRichRunsCore(ScaleTableRunsForShrink(cell.Runs, runFontSizeScale, minimumShrinkFontSize), wrapWidth, fontSize, baseFont, leading, null, DefaultParagraphTabStopWidth, options);
        if (wrap.Lines.Count == 0) {
            wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
        }

        while (wrap.LineHeights.Count < wrap.Lines.Count) {
            wrap.LineHeights.Add(leading);
        }

        return new TableCellTextLayout(wrap.Lines, wrap.LineHeights);
    }

    private static System.Collections.Generic.IReadOnlyList<TextRun> ScaleTableRunsForShrink(System.Collections.Generic.IReadOnlyList<TextRun> runs, double runFontSizeScale, double minimumShrinkFontSize) {
        if (runFontSizeScale >= 0.999D) {
            return runs;
        }

        double minimumExplicitFontSize = minimumShrinkFontSize > 0D ? minimumShrinkFontSize : 0.001D;
        var scaledRuns = new System.Collections.Generic.List<TextRun>(runs.Count);
        foreach (TextRun run in runs) {
            if (run.InlineElement != null) {
                scaledRuns.Add(run);
                continue;
            }

            double? scaledFontSize = null;
            if (run.FontSize.HasValue) {
                scaledFontSize = run.FontSize.Value <= minimumExplicitFontSize
                    ? run.FontSize.Value
                    : System.Math.Max(minimumExplicitFontSize, run.FontSize.Value * runFontSizeScale);
            }

            scaledRuns.Add(new TextRun(
                run.Text,
                run.Bold,
                run.Underline,
                run.Color,
                run.Italic,
                run.Strike,
                scaledFontSize,
                run.Font,
                run.LinkUri,
                run.LinkContents,
                run.Baseline,
                run.LinkDestinationName,
                run.TabLeader,
                run.TabAlignment,
                run.BackgroundColor,
                run.FontFamily));
        }

        return scaledRuns.AsReadOnly();
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> ScaleTableCellParagraphsForShrink(System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> paragraphs, double runFontSizeScale, double minimumShrinkFontSize) {
        if (runFontSizeScale >= 0.999D) {
            return paragraphs;
        }

        var scaledParagraphs = new System.Collections.Generic.List<PdfTableCellParagraph>(paragraphs.Count);
        foreach (PdfTableCellParagraph paragraph in paragraphs) {
            scaledParagraphs.Add(new PdfTableCellParagraph(
                ScaleTableRunsForShrink(paragraph.Runs, runFontSizeScale, minimumShrinkFontSize),
                paragraph.SpacingAfter,
                paragraph.Align,
                paragraph.SpacingBefore,
                paragraph.LeftIndent,
                paragraph.RightIndent,
                paragraph.FirstLineIndent,
                paragraph.LineHeight,
                paragraph.DefaultTabStopWidth,
                paragraph.TabStops));
        }

        return scaledParagraphs.AsReadOnly();
    }

    private static double GetTableCellWrapWidth(double innerWidth, bool noWrap) =>
        noWrap ? TableCellNoWrapWidth : innerWidth;

    private static TableCellTextLayout CreateTableCellParagraphTextLayout(System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> paragraphs, double wrapWidth, double cellInnerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
        var lineHeights = new System.Collections.Generic.List<double>();
        var lineAlignments = new System.Collections.Generic.List<PdfAlign?>();
        var lineXOffsets = new System.Collections.Generic.List<double>();
        var lineWidths = new System.Collections.Generic.List<double>();
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            PdfTableCellParagraph paragraph = paragraphs[paragraphIndex];
            PdfParagraphStyle paragraphStyle = CreateTableCellParagraphStyle(paragraph, cellInnerWidth);
            double paragraphLeading = paragraphStyle.LineHeight.HasValue ? GetParagraphLeading(paragraphStyle, fontSize) : leading;
            var paragraphFrame = GetParagraphTextFrame(paragraphStyle, 0D, wrapWidth);
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(
                paragraph.Runs,
                paragraphFrame.Width,
                fontSize,
                baseFont,
                paragraphLeading,
                paragraphFrame.FirstLineWidth,
                paragraphFrame.FirstLineX - paragraphFrame.X,
                GetParagraphTabStopWidth(paragraphStyle),
                options,
                paragraphStyle.TabStops.ToArray());
            if (wrap.Lines.Count == 0) {
                wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
            }

            while (wrap.LineHeights.Count < wrap.Lines.Count) {
                wrap.LineHeights.Add(paragraphLeading);
            }

            int firstNewLineIndex = lines.Count;
            if (paragraph.SpacingBefore > 0D && lineHeights.Count > 0) {
                lineHeights[lineHeights.Count - 1] += paragraph.SpacingBefore;
            }

            lines.AddRange(wrap.Lines);
            lineHeights.AddRange(wrap.LineHeights);
            for (int lineIndex = firstNewLineIndex; lineIndex < lines.Count; lineIndex++) {
                lineAlignments.Add(paragraph.Align);
                bool firstParagraphLine = lineIndex == firstNewLineIndex;
                lineXOffsets.Add(firstParagraphLine ? paragraphFrame.FirstLineX : paragraphFrame.X);
                lineWidths.Add(firstParagraphLine ? paragraphFrame.FirstLineWidth : paragraphFrame.Width);
            }

            if (paragraphIndex < paragraphs.Count - 1 && lines.Count > firstNewLineIndex) {
                MarkRichLineTextSeparator(lines[lines.Count - 1]);
            }

            if (paragraph.SpacingAfter > 0D && lineHeights.Count > firstNewLineIndex) {
                int lastParagraphLineIndex = lineHeights.Count - 1;
                lineHeights[lastParagraphLineIndex] += paragraph.SpacingAfter;
            }
        }

        if (lines.Count == 0) {
            lines.Add(new System.Collections.Generic.List<RichSeg>());
            lineHeights.Add(leading);
            lineAlignments.Add(null);
            lineXOffsets.Add(0D);
            lineWidths.Add(wrapWidth);
        }

        return new TableCellTextLayout(lines, lineHeights, lineAlignments, lineXOffsets, lineWidths);
    }

    private static PdfParagraphStyle CreateTableCellParagraphStyle(PdfTableCellParagraph paragraph, double availableWidth) {
        const double minimumTextWidth = 0.001D;
        double safeWidth = double.IsNaN(availableWidth) || double.IsInfinity(availableWidth)
            ? minimumTextWidth
            : System.Math.Max(minimumTextWidth, availableWidth);
        double leftIndent = System.Math.Min(paragraph.LeftIndent, System.Math.Max(0D, safeWidth - minimumTextWidth));
        double rightIndent = System.Math.Min(paragraph.RightIndent, System.Math.Max(0D, safeWidth - leftIndent - minimumTextWidth));
        double textWidth = System.Math.Max(minimumTextWidth, safeWidth - leftIndent - rightIndent);
        double firstLineIndent = System.Math.Max(-leftIndent, System.Math.Min(paragraph.FirstLineIndent, textWidth - minimumTextWidth));
        var style = new PdfParagraphStyle {
            LineHeight = paragraph.LineHeight,
            LeftIndent = leftIndent,
            RightIndent = rightIndent,
            FirstLineIndent = firstLineIndent,
            DefaultTabStopWidth = paragraph.DefaultTabStopWidth
        };

        foreach (PdfTabStop tabStop in paragraph.TabStops) {
            style.TabStops.Add(tabStop.Clone());
        }

        return style;
    }

    private static TableCellTextLayout CreateListItemTextLayout(PdfListItem item, double innerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options) {
        var wrap = WrapRichRunsCore(item.Runs, innerWidth, fontSize, baseFont, leading, null, DefaultParagraphTabStopWidth, options);
        if (wrap.Lines.Count == 0) {
            wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
        }

        while (wrap.LineHeights.Count < wrap.Lines.Count) {
            wrap.LineHeights.Add(leading);
        }

        return new TableCellTextLayout(wrap.Lines, wrap.LineHeights);
    }

    private static double GetRichLineHeight(System.Collections.Generic.IReadOnlyList<double> heights, int lineIndex, double fallbackLeading) =>
        lineIndex >= 0 && lineIndex < heights.Count ? heights[lineIndex] : fallbackLeading;

    private static int LimitTableCellLineCountToHeight(TableCellTextLayout lines, int startLine, int requestedLineCount, double fallbackLeading, double availableHeight) {
        int maximumLineCount = System.Math.Max(0, System.Math.Min(requestedLineCount, lines.LineCount - startLine));
        double consumedHeight = 0D;
        int visibleLineCount = 0;
        for (int offset = 0; offset < maximumLineCount; offset++) {
            double lineHeight = GetRichLineHeight(lines.LineHeights, startLine + offset, fallbackLeading);
            if (consumedHeight + lineHeight > availableHeight + 0.001D) {
                break;
            }

            consumedHeight += lineHeight;
            visibleLineCount++;
        }

        return visibleLineCount;
    }

    private static double MeasureRichLinesHeight(System.Collections.Generic.IReadOnlyList<double> heights, int lineCount, double fallbackLeading) {
        double height = 0D;
        for (int index = 0; index < lineCount; index++) {
            height += GetRichLineHeight(heights, index, fallbackLeading);
        }

        return height;
    }

    private static double MeasureTableCellTextHeight(TableCellTextLayout layout, int startLine, int lineCount, double fallbackLeading) {
        int available = System.Math.Max(0, layout.Lines.Count - startLine);
        int visible = System.Math.Max(0, System.Math.Min(lineCount, available));
        if (visible == 0) {
            return fallbackLeading;
        }

        double height = 0D;
        for (int i = 0; i < visible; i++) {
            int lineIndex = startLine + i;
            height += lineIndex < layout.LineHeights.Count ? layout.LineHeights[lineIndex] : fallbackLeading;
        }

        return height;
    }

    private static (double Width, double Height) ResolveTableCellImageBox(PdfTableCellImage image, double innerWidth) {
        if (image.Style?.ScaleDownToFit != true || innerWidth <= 0D) {
            return (image.Width, image.Height);
        }

        double scale = Math.Min(1D, innerWidth / image.Width);
        if (scale <= 0D || double.IsNaN(scale) || double.IsInfinity(scale)) {
            return (image.Width, image.Height);
        }

        return (image.Width * scale, image.Height * scale);
    }

    private static double MeasureTableCellObjectStackHeight(TableCellLayout cell, double innerWidth) {
        if (cell.Images.Count == 0 && cell.CheckBoxes.Count == 0 && cell.FormFields.Count == 0) {
            return 0D;
        }

        double height = 0D;
        int objectCount = 0;
        for (int index = 0; index < cell.Images.Count; index++) {
            if (objectCount > 0) {
                height += TableCellCheckBoxGap;
            }

            height += ResolveTableCellImageBox(cell.Images[index], innerWidth).Height;
            objectCount++;
        }

        for (int index = 0; index < cell.CheckBoxes.Count; index++) {
            if (objectCount > 0) {
                height += TableCellCheckBoxGap;
            }

            height += cell.CheckBoxes[index].Size;
            objectCount++;
        }

        for (int index = 0; index < cell.FormFields.Count; index++) {
            if (objectCount > 0) {
                height += TableCellCheckBoxGap;
            }

            height += cell.FormFields[index].Height;
            objectCount++;
        }

        return height;
    }

    private static double MeasureTableCellContentHeight(TableCellLayout cell, TableCellTextLayout layout, int startLine, int lineCount, double fallbackLeading, double innerWidth) =>
        MeasureTableCellContentHeight(cell, layout, startLine, lineCount, fallbackLeading, innerWidth, includeObjects: true);

    private static double MeasureTableCellContentHeight(TableCellLayout cell, TableCellTextLayout layout, int startLine, int lineCount, double fallbackLeading, double innerWidth, bool includeObjects) {
        double textHeight = MeasureTableCellTextHeight(layout, startLine, lineCount, fallbackLeading);
        if (!includeObjects) {
            return textHeight;
        }

        double objectStackHeight = MeasureTableCellObjectStackHeight(cell, innerWidth);
        if (objectStackHeight <= 0D) {
            return textHeight;
        }

        if (CanRenderTableCellCheckBoxInline(cell, layout, startLine, lineCount)) {
            return System.Math.Max(textHeight, cell.CheckBoxes[0].Size);
        }

        if (string.IsNullOrEmpty(cell.Text)) {
            return objectStackHeight;
        }

        return textHeight + TableCellCheckBoxGap + objectStackHeight;
    }

    private static double MeasureTableCellObjectWidth(TableCellLayout cell) {
        double width = 0D;
        for (int index = 0; index < cell.Images.Count; index++) {
            PdfTableCellImage image = cell.Images[index];
            width = System.Math.Max(width, image.Style?.ScaleDownToFit == true ? 1D : image.Width);
        }

        for (int index = 0; index < cell.CheckBoxes.Count; index++) {
            width = System.Math.Max(width, cell.CheckBoxes[index].Size);
        }

        for (int index = 0; index < cell.FormFields.Count; index++) {
            width = System.Math.Max(width, cell.FormFields[index].Width);
        }

        return width;
    }

    private static void ValidateTableCellTextWidths(TableBlock table, PdfTableStyle style, int columnCount, double[] columnWidths, double columnGap) {
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double padLeft = GetTableCellPaddingLeft(style, rowIndex, cell.Column);
                double padRight = GetTableCellPaddingRight(style, rowIndex, cell.Column);
                if (cellWidth - padLeft - padRight <= 0.001) {
                    throw new ArgumentException("Table horizontal cell padding must leave a positive text width.");
                }
            }
        }
    }

    private static void ValidateTableRowStyleBounds(PdfTableStyle style, int rowCount) {
        if (style.RowMinHeights != null) {
            for (int row = rowCount; row < style.RowMinHeights.Count; row++) {
                if (style.RowMinHeights[row].HasValue) {
                    throw new ArgumentException("Table row minimum heights must fit inside the table grid.");
                }
            }
        }

        if (style.FixedRowHeights != null) {
            for (int row = rowCount; row < style.FixedRowHeights.Count; row++) {
                if (style.FixedRowHeights[row].HasValue) {
                    throw new ArgumentException("Table fixed row heights must fit inside the table grid.");
                }
            }
        }

        if (style.RowAllowBreakAcrossPages != null) {
            for (int row = rowCount; row < style.RowAllowBreakAcrossPages.Count; row++) {
                if (style.RowAllowBreakAcrossPages[row].HasValue) {
                    throw new ArgumentException("Table row break policies must fit inside the table grid.");
                }
            }
        }
    }

    private static void ApplyTableRowSpanHeights(TableBlock table, PdfTableStyle style, int columnCount, double[] columnWidths, TableCellTextLayout[][] rowLines, double[] rowHeights, double[] rowLeadings, double columnGap, double rowGap) {
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1) {
                    continue;
                }

                int rowSpan = System.Math.Min(cell.RowSpan, rowHeights.Length - rowIndex);
                if (rowSpan <= 1) {
                    continue;
                }

                TableCellTextLayout lines = rowLines[rowIndex][cell.Column];
                double cellWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double innerWidth = Math.Max(1D, cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) - GetTableCellPaddingRight(style, rowIndex, cell.Column));
                double requiredHeight = MeasureTableCellContentHeight(cell, lines, 0, lines.LineCount, rowLeadings[rowIndex], innerWidth) +
                    GetTableCellPaddingTop(style, rowIndex, cell.Column) +
                    GetTableCellPaddingBottom(style, rowIndex, cell.Column);
                double currentHeight = GetTableCellHeight(rowHeights, rowIndex, rowSpan, rowGap);
                if (requiredHeight <= currentHeight + 0.001) {
                    continue;
                }

                var flexibleRows = new System.Collections.Generic.List<int>(rowSpan);
                for (int spanRow = rowIndex; spanRow < rowIndex + rowSpan; spanRow++) {
                    if (!GetTableRowFixedHeight(style, spanRow).HasValue) {
                        flexibleRows.Add(spanRow);
                    }
                }

                if (flexibleRows.Count == 0) {
                    continue;
                }

                double extraPerRow = (requiredHeight - currentHeight) / flexibleRows.Count;
                for (int index = 0; index < flexibleRows.Count; index++) {
                    int spanRow = flexibleRows[index];
                    rowHeights[spanRow] += extraPerRow;
                }
            }
        }
    }

    private static void ValidateTableRowSpansWithinRoleBoundaries(TableBlock table, int columnCount, int headerRowCount, int footerStartRowIndex) {
        for (int rowIndex = 0; rowIndex < table.Cells.Count; rowIndex++) {
            var cells = GetTableCellLayouts(table, rowIndex, columnCount);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (cell.RowSpan <= 1) {
                    continue;
                }

                int lastRowExclusive = rowIndex + cell.RowSpan;
                if (rowIndex < headerRowCount && lastRowExclusive > headerRowCount) {
                    throw new ArgumentException("Table cell row span cannot cross the table header boundary.");
                }

                if (rowIndex < footerStartRowIndex && lastRowExclusive > footerStartRowIndex) {
                    throw new ArgumentException("Table cell row span cannot cross the table footer boundary.");
                }
            }
        }
    }

    private static bool TryGetTableCellLayoutAtColumn(System.Collections.Generic.List<TableCellLayout> cells, int column, out TableCellLayout layout) {
        for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
            if (cells[cellIndex].Column == column) {
                layout = cells[cellIndex];
                return true;
            }
        }

        layout = default;
        return false;
    }

    private static bool IsTableBoundaryInsideSpannedCell(TableBlock table, int rowIndex, int boundaryColumn, int columnCount) {
        if (rowIndex < 0 || rowIndex >= table.Cells.Count || boundaryColumn < 0 || boundaryColumn >= columnCount - 1) {
            return false;
        }

        for (int sourceRowIndex = 0; sourceRowIndex <= rowIndex; sourceRowIndex++) {
            var cells = GetTableCellLayouts(table, sourceRowIndex, columnCount);
            for (int i = 0; i < cells.Count; i++) {
                TableCellLayout cell = cells[i];
                if (sourceRowIndex + cell.RowSpan <= rowIndex) {
                    continue;
                }

                if (cell.Column <= boundaryColumn && boundaryColumn < cell.Column + cell.ColumnSpan - 1) {
                    return true;
                }
            }
        }

        return false;
    }

}
