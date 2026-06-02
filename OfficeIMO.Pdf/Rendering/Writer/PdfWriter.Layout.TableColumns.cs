using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double? GetOptionalColumnWidth(System.Collections.Generic.List<double?>? values, int index, string errorMessage) {
        if (values == null || index >= values.Count || !values[index].HasValue) {
            return null;
        }

        double value = values[index]!.Value;
        if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentException(errorMessage);
        }

        return value;
    }

    private static double ResolveTableFrameWidth(PdfTableStyle style, double containerWidth) {
        if (style.LeftIndent < 0 || double.IsNaN(style.LeftIndent) || double.IsInfinity(style.LeftIndent)) {
            throw new ArgumentException("Table left indent must be a non-negative finite value.");
        }

        double frameWidth = containerWidth - style.LeftIndent;
        if (frameWidth <= 0.001 || double.IsNaN(frameWidth) || double.IsInfinity(frameWidth)) {
            throw new ArgumentException("Table left indent must leave a positive table width.");
        }

        return frameWidth;
    }

    private static double ResolveTableAvailableWidth(PdfTableStyle style, double containerWidth) {
        double frameWidth = ResolveTableFrameWidth(style, containerWidth);
        if (style.MaxWidth.HasValue) {
            double maxWidth = style.MaxWidth.Value;
            if (maxWidth <= 0 || double.IsNaN(maxWidth) || double.IsInfinity(maxWidth)) {
                throw new ArgumentException("Table max width must be a positive finite value.");
            }

            return Math.Min(frameWidth, maxWidth);
        }

        return frameWidth;
    }

    private static double FitFixedTableColumnsToAvailableWidth(double[] columnWidths, bool[] fixedColumns, double?[] minWidths, double fixedWidthTotal, double availableWidth) {
        if (fixedWidthTotal <= availableWidth + 0.001D) {
            return fixedWidthTotal;
        }

        double requiredMinimumWidth = 0D;
        for (int column = 0; column < columnWidths.Length; column++) {
            if (fixedColumns[column] && minWidths[column].HasValue) {
                requiredMinimumWidth += minWidths[column]!.Value;
            }
        }

        if (requiredMinimumWidth > availableWidth + 0.001D) {
            throw new ArgumentException("Table fixed column widths cannot fit inside the available table width after applying minimum widths.");
        }

        double[] originalWidths = new double[columnWidths.Length];
        bool[] lockedColumns = new bool[columnWidths.Length];
        double remainingOriginalWidth = 0D;
        double remainingAvailableWidth = availableWidth;
        for (int column = 0; column < columnWidths.Length; column++) {
            if (!fixedColumns[column]) {
                continue;
            }

            originalWidths[column] = columnWidths[column];
            remainingOriginalWidth += columnWidths[column];
        }

        while (remainingOriginalWidth > 0.001D) {
            double scale = remainingAvailableWidth / remainingOriginalWidth;
            bool lockedMinimum = false;

            for (int column = 0; column < columnWidths.Length; column++) {
                if (!fixedColumns[column] || lockedColumns[column]) {
                    continue;
                }

                double candidateWidth = originalWidths[column] * scale;
                if (minWidths[column].HasValue && candidateWidth < minWidths[column]!.Value - 0.001D) {
                    columnWidths[column] = minWidths[column]!.Value;
                    lockedColumns[column] = true;
                    remainingAvailableWidth -= columnWidths[column];
                    remainingOriginalWidth -= originalWidths[column];
                    lockedMinimum = true;
                }
            }

            if (!lockedMinimum) {
                for (int column = 0; column < columnWidths.Length; column++) {
                    if (fixedColumns[column] && !lockedColumns[column]) {
                        columnWidths[column] = originalWidths[column] * scale;
                    }
                }

                break;
            }
        }

        return fixedColumns.Select((fixedColumn, column) => fixedColumn ? columnWidths[column] : 0D).Sum();
    }

    private static TableColumnLayout ResolveTableColumnLayout(TableBlock table, PdfOptions options, PdfTableStyle style, int columns, double frameWidth, double fontSize, int headerRowCount, int footerStartRowIndex) {
        double[]? autoFitWeights = style.AutoFitColumns
            ? MeasureAutoFitColumnWeights(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double[]? autoFitMinimumWidths = style.AutoFitColumns
            ? MeasureAutoFitColumnMinimumWidths(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double columnGap = GetTableCellSpacing(style);
        double tableWidth = ResolveTableAvailableWidth(style, frameWidth);
        double tableInnerWidth = tableWidth - (columns - 1) * columnGap;
        if (tableInnerWidth <= 0.001 || double.IsNaN(tableInnerWidth) || double.IsInfinity(tableInnerWidth)) {
            throw new ArgumentException("Table cell spacing must leave a positive table width.");
        }

        double[] columnWidths = new double[columns];
        double[] columnWeights = new double[columns];
        bool[] fixedColumns = new bool[columns];
        double?[] minWidths = new double?[columns];
        double?[] maxWidths = new double?[columns];
        double fixedWidthTotal = 0D;
        double totalWeight = 0D;

        for (int column = 0; column < columns; column++) {
            double? minWidth = GetOptionalColumnWidth(style.ColumnMinWidthPoints, column, "Table minimum column widths must be positive finite values.");
            if (!minWidth.HasValue && autoFitMinimumWidths != null && column < autoFitMinimumWidths.Length) {
                minWidth = autoFitMinimumWidths[column];
            }

            double? maxWidth = GetOptionalColumnWidth(style.ColumnMaxWidthPoints, column, "Table maximum column widths must be positive finite values.");
            if (minWidth.HasValue && maxWidth.HasValue && minWidth.Value > maxWidth.Value + 0.001) {
                throw new ArgumentException("Table minimum column widths cannot be greater than maximum column widths.");
            }

            minWidths[column] = minWidth;
            maxWidths[column] = maxWidth;

            if (style.ColumnWidthPoints != null &&
                column < style.ColumnWidthPoints.Count &&
                style.ColumnWidthPoints[column].HasValue) {
                double fixedWidth = style.ColumnWidthPoints[column]!.Value;
                if (fixedWidth <= 0 || double.IsNaN(fixedWidth) || double.IsInfinity(fixedWidth)) {
                    throw new ArgumentException("Table fixed column widths must be positive finite values.");
                }

                if (minWidth.HasValue && fixedWidth < minWidth.Value - 0.001) {
                    throw new ArgumentException("Table fixed column widths cannot be smaller than configured minimum widths.");
                }

                if (maxWidth.HasValue && fixedWidth > maxWidth.Value + 0.001) {
                    throw new ArgumentException("Table fixed column widths cannot be greater than configured maximum widths.");
                }

                columnWidths[column] = fixedWidth;
                fixedColumns[column] = true;
                fixedWidthTotal += fixedWidth;
                continue;
            }

            double weight = 1D;
            if (style.ColumnWidthWeights != null && column < style.ColumnWidthWeights.Count) {
                weight = style.ColumnWidthWeights[column];
                if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                    throw new ArgumentException("Table column width weights must be positive finite values.");
                }
            } else if (autoFitWeights != null && column < autoFitWeights.Length) {
                weight = autoFitWeights[column];
            }

            columnWeights[column] = weight;
            totalWeight += weight;
        }

        fixedWidthTotal = FitFixedTableColumnsToAvailableWidth(columnWidths, fixedColumns, minWidths, fixedWidthTotal, tableInnerWidth);

        double remainingWidth = Math.Max(0D, tableInnerWidth - fixedWidthTotal);
        if (totalWeight <= 0D) {
            remainingWidth = 0D;
        }

        DistributeFlexibleColumns(columnWidths, columnWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
        tableWidth = Math.Min(tableWidth, columnWidths.Sum() + (columns - 1) * columnGap);
        ValidateTableCellTextWidths(table, style, columns, columnWidths, columnGap);

        return new TableColumnLayout {
            Widths = columnWidths,
            Width = tableWidth
        };
    }

    private static double ResolveTableX(PdfAlign align, PdfTableStyle style, double containerLeft, double containerWidth, double tableWidth) {
        double frameLeft = containerLeft + style.LeftIndent;
        double frameWidth = ResolveTableFrameWidth(style, containerWidth);
        if (align == PdfAlign.Center) {
            return frameLeft + Math.Max(0, (frameWidth - tableWidth) / 2);
        }

        if (align == PdfAlign.Right) {
            return frameLeft + Math.Max(0, frameWidth - tableWidth);
        }

        return frameLeft;
    }

    private static bool IsValidPdfAlign(PdfAlign align) =>
        align == PdfAlign.Left || align == PdfAlign.Center || align == PdfAlign.Right;

    private static bool IsValidPdfColumnAlign(PdfColumnAlign align) =>
        align == PdfColumnAlign.Left || align == PdfColumnAlign.Center || align == PdfColumnAlign.Right;

    private static bool IsValidPdfCellVerticalAlign(PdfCellVerticalAlign align) =>
        align == PdfCellVerticalAlign.Top || align == PdfCellVerticalAlign.Middle || align == PdfCellVerticalAlign.Bottom;

    private static OfficeIMO.Drawing.OfficeFontInfo ToOfficeFontInfo(PdfStandardFont font, double size) {
        string family = font switch {
            PdfStandardFont.TimesRoman or PdfStandardFont.TimesBold or PdfStandardFont.TimesItalic or PdfStandardFont.TimesBoldItalic => "Times New Roman",
            PdfStandardFont.Courier or PdfStandardFont.CourierBold or PdfStandardFont.CourierOblique or PdfStandardFont.CourierBoldOblique => "Courier New",
            _ => "Helvetica"
        };

        OfficeIMO.Drawing.OfficeFontStyle style = OfficeIMO.Drawing.OfficeFontStyle.Regular;
        switch (font) {
            case PdfStandardFont.HelveticaBold:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesBold:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierBold:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Bold;
                break;
        }

        switch (font) {
            case PdfStandardFont.HelveticaOblique:
            case PdfStandardFont.HelveticaBoldOblique:
            case PdfStandardFont.TimesItalic:
            case PdfStandardFont.TimesBoldItalic:
            case PdfStandardFont.CourierOblique:
            case PdfStandardFont.CourierBoldOblique:
                style |= OfficeIMO.Drawing.OfficeFontStyle.Italic;
                break;
        }

        return new OfficeIMO.Drawing.OfficeFontInfo(family, size, style);
    }

    private static double[] MeasureAutoFitColumnWeights(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var weights = new double[cols];
        var normalFont = ToOfficeFontInfo(ChooseNormal(options.DefaultFont), fontSize);
        var measurer = OfficeIMO.Drawing.OfficeTextMeasurer.Create(normalFont);

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = ToOfficeFontInfo(GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex)), rowSize);
            var measurementStyle = measurer.CreateStyle(rowFont);
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double measuredPoints = System.Math.Max(
                    measurer.MeasureWidth(cell.Text, measurementStyle) * 72D / measurementStyle.Dpi,
                    MeasureTableCellObjectWidth(cell));
                double requestedWidth = Math.Max(1D, measuredPoints + GetTableCellPaddingLeft(style, rowIndex, cell.Column) + GetTableCellPaddingRight(style, rowIndex, cell.Column));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int c = cell.Column; c < cell.Column + cell.ColumnSpan && c < cols; c++) {
                    if (requestedPerColumn > weights[c]) {
                        weights[c] = requestedPerColumn;
                    }
                }
            }
        }

        for (int c = 0; c < weights.Length; c++) {
            if (weights[c] <= 0D) {
                weights[c] = 1D;
            }
        }

        return weights;
    }

    private static double[] MeasureAutoFitColumnMinimumWidths(TableBlock table, PdfOptions options, PdfTableStyle style, double fontSize, int headerRowCount, int footerStartRowIndex) {
        int cols = GetTableColumnCount(table);
        var widths = new double[cols];
        double maximumTokenWidth = Math.Max(1D, fontSize * Math.Max(4D, 13D - cols));

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            double rowSize = GetTableRowFontSize(style, rowIndex, headerRowCount, footerStartRowIndex, fontSize);
            var rowFont = GetTableRowFont(options, GetTableRowBold(style, rowIndex, headerRowCount, footerStartRowIndex));
            var cells = GetTableCellLayouts(table, rowIndex, cols);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                double tokenWidth = 0D;
                string[] tokens = cell.Text
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n')
                    .Split(TokenSplitChars, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length == 0) {
                    tokenWidth = EstimateSimpleTextWidth(cell.Text, rowFont, rowSize);
                } else {
                    for (int tokenIndex = 0; tokenIndex < tokens.Length; tokenIndex++) {
                        tokenWidth = Math.Max(tokenWidth, EstimateSimpleTextWidth(tokens[tokenIndex], rowFont, rowSize));
                    }
                }

                double requestedWidth = Math.Max(1D, System.Math.Max(Math.Min(tokenWidth, maximumTokenWidth), MeasureTableCellObjectWidth(cell)) + GetTableCellPaddingLeft(style, rowIndex, cell.Column) + GetTableCellPaddingRight(style, rowIndex, cell.Column));
                double requestedPerColumn = requestedWidth / cell.ColumnSpan;
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < cols; columnIndex++) {
                    if (requestedPerColumn > widths[columnIndex]) {
                        widths[columnIndex] = requestedPerColumn;
                    }
                }
            }
        }

        for (int columnIndex = 0; columnIndex < widths.Length; columnIndex++) {
            if (widths[columnIndex] <= 0D) {
                widths[columnIndex] = 1D;
            }
        }

        return widths;
    }

    private static void DistributeFlexibleColumns(
        double[] widths,
        double[] weights,
        bool[] fixedColumns,
        double?[] minWidths,
        double?[] maxWidths,
        double availableWidth) {
        var active = new bool[widths.Length];
        int activeCount = 0;
        double requiredMinimum = 0;

        for (int i = 0; i < widths.Length; i++) {
            if (fixedColumns[i]) {
                continue;
            }

            active[i] = true;
            activeCount++;
            if (minWidths[i].HasValue) {
                requiredMinimum += minWidths[i]!.Value;
            }
        }

        if (requiredMinimum > availableWidth + 0.001) {
            throw new ArgumentException("Table minimum column widths exceed the available table width.");
        }

        double remaining = availableWidth;
        while (activeCount > 0) {
            double weightSum = 0;
            for (int i = 0; i < weights.Length; i++) {
                if (active[i]) {
                    weightSum += weights[i];
                }
            }

            bool constrained = false;
            for (int i = 0; i < widths.Length; i++) {
                if (!active[i]) {
                    continue;
                }

                double proposed = weightSum > 0 ? remaining * (weights[i] / weightSum) : remaining / activeCount;
                if (minWidths[i].HasValue && proposed < minWidths[i]!.Value) {
                    widths[i] = minWidths[i]!.Value;
                    remaining -= widths[i];
                    active[i] = false;
                    activeCount--;
                    constrained = true;
                } else if (maxWidths[i].HasValue && proposed > maxWidths[i]!.Value) {
                    widths[i] = maxWidths[i]!.Value;
                    remaining -= widths[i];
                    active[i] = false;
                    activeCount--;
                    constrained = true;
                }
            }

            if (constrained) {
                continue;
            }

            for (int i = 0; i < widths.Length; i++) {
                if (!active[i]) {
                    continue;
                }

                widths[i] = weightSum > 0 ? remaining * (weights[i] / weightSum) : remaining / activeCount;
                active[i] = false;
            }
            break;
        }
    }


}
