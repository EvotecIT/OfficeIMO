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

    private static double ResolveTableLayoutWidth(
        PdfTableStyle style,
        double containerWidth,
        double[]? autoFitPreferredWidths,
        double[]? autoFitMinimumWidths,
        int columnCount,
        double columnGap) {
        double availableWidth = ResolveTableAvailableWidth(style, containerWidth);
        if (!style.PreferredWidth.HasValue) {
            return availableWidth;
        }

        double preferredWidth = Math.Min(availableWidth, style.PreferredWidth.Value);
        double measuredContentWidth = 0D;
        if (autoFitPreferredWidths != null && autoFitPreferredWidths.Length > 0) {
            measuredContentWidth = Math.Max(measuredContentWidth, autoFitPreferredWidths.Sum());
        }

        if (autoFitMinimumWidths != null && autoFitMinimumWidths.Length > 0) {
            measuredContentWidth = Math.Max(measuredContentWidth, autoFitMinimumWidths.Sum());
        }

        measuredContentWidth += Math.Max(0, columnCount - 1) * columnGap;
        return Math.Min(availableWidth, Math.Max(preferredWidth, measuredContentWidth));
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

    private static double ExpandFixedTableColumnsToAvailableWidth(double[] columnWidths, bool[] fixedColumns, double?[] maxWidths, double fixedWidthTotal, double availableWidth) {
        if (fixedWidthTotal <= 0D || fixedWidthTotal >= availableWidth - 0.001D) {
            return fixedWidthTotal;
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
            bool lockedMaximum = false;

            for (int column = 0; column < columnWidths.Length; column++) {
                if (!fixedColumns[column] || lockedColumns[column]) {
                    continue;
                }

                double candidateWidth = originalWidths[column] * scale;
                if (maxWidths[column].HasValue && candidateWidth > maxWidths[column]!.Value + 0.001D) {
                    columnWidths[column] = maxWidths[column]!.Value;
                    lockedColumns[column] = true;
                    remainingAvailableWidth -= columnWidths[column];
                    remainingOriginalWidth -= originalWidths[column];
                    lockedMaximum = true;
                }
            }

            if (!lockedMaximum) {
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
        AutoFitColumnProfile[]? autoFitProfiles = style.AutoFitColumns
            ? MeasureAutoFitColumnProfiles(table, headerRowCount)
            : null;
        double[]? autoFitWeights = style.AutoFitColumns
            ? MeasureAutoFitColumnWeights(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double[]? autoFitMinimumWidths = style.AutoFitColumns
            ? MeasureAutoFitColumnMinimumWidths(table, options, style, fontSize, headerRowCount, footerStartRowIndex)
            : null;
        double columnGap = GetTableCellSpacing(style);
        double tableWidth = ResolveTableLayoutWidth(style, frameWidth, autoFitWeights, autoFitMinimumWidths, columns, columnGap);
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

            bool hasExplicitWeight = style.ColumnWidthWeights != null && column < style.ColumnWidthWeights.Count;
            double weight = 1D;
            if (hasExplicitWeight) {
                weight = style.ColumnWidthWeights![column];
                if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                    throw new ArgumentException("Table column width weights must be positive finite values.");
                }
            } else if (autoFitWeights != null && column < autoFitWeights.Length) {
                weight = autoFitWeights[column];
            }

            if (!hasExplicitWeight && autoFitWeights != null && minWidth.HasValue) {
                AutoFitColumnProfile profile = autoFitProfiles != null && column < autoFitProfiles.Length
                    ? autoFitProfiles[column]
                    : default;
                weight = ResolveAutoFitFlexibleWeight(weight, minWidth.Value, profile);
            }

            columnWeights[column] = weight;
            totalWeight += weight;
        }

        fixedWidthTotal = FitFixedTableColumnsToAvailableWidth(columnWidths, fixedColumns, minWidths, fixedWidthTotal, tableInnerWidth);
        if (style.PreserveWidth && totalWeight <= 0D) {
            fixedWidthTotal = ExpandFixedTableColumnsToAvailableWidth(columnWidths, fixedColumns, maxWidths, fixedWidthTotal, tableInnerWidth);
        }

        double remainingWidth = Math.Max(0D, tableInnerWidth - fixedWidthTotal);
        if (totalWeight <= 0D) {
            remainingWidth = 0D;
        }

        DistributeFlexibleColumns(columnWidths, columnWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
        if (!style.PreserveWidth) {
            tableWidth = Math.Min(tableWidth, columnWidths.Sum() + (columns - 1) * columnGap);
        }
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
            double scale = availableWidth > 0D ? availableWidth / requiredMinimum : 0D;
            for (int i = 0; i < widths.Length; i++) {
                if (!active[i]) {
                    continue;
                }

                widths[i] = (minWidths[i] ?? 0D) * scale;
            }

            return;
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
