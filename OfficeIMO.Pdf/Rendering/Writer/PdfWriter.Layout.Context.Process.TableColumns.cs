using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private sealed class PreparedTableColumns {
            public PreparedTableColumns(double tableWidth, double[] columnWidths) {
                TableWidth = tableWidth;
                ColumnWidths = columnWidths;
            }

            public double TableWidth { get; }
            public double[] ColumnWidths { get; }
        }

        private PreparedTableColumns PrepareTableColumns(TableBlock table, PdfTableStyle style, double availableWidth, double fontSize, int headerRowCount, int footerStartRowIndex) {
            int columns = GetTableColumnCount(table);
            double columnGap = GetTableCellSpacing(style);
            AutoFitColumnProfile[]? autoFitProfiles = style.AutoFitColumns
                ? MeasureAutoFitColumnProfiles(table, headerRowCount)
                : null;
            double[]? autoFitWeights = style.AutoFitColumns
                ? MeasureAutoFitColumnWeights(table, currentOpts, style, fontSize, headerRowCount, footerStartRowIndex)
                : null;
            double[]? autoFitMinimumWidths = style.AutoFitColumns
                ? MeasureAutoFitColumnMinimumWidths(table, currentOpts, style, fontSize, headerRowCount, footerStartRowIndex)
                : null;
            double tableWidth = ResolveTableLayoutWidth(style, availableWidth, autoFitWeights, autoFitMinimumWidths, columns, columnGap);
            double[] columnWidths = new double[columns];
            double[] columnWeights = new double[columns];
            bool[] fixedColumns = new bool[columns];
            double?[] minWidths = new double?[columns];
            double?[] maxWidths = new double?[columns];
            double fixedWidthTotal = 0;
            double totalWeight = 0;

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

                double weight = 1;
                if (style.ColumnWidthWeights != null && column < style.ColumnWidthWeights.Count) {
                    weight = style.ColumnWidthWeights[column];
                    if (weight <= 0 || double.IsNaN(weight) || double.IsInfinity(weight)) {
                        throw new ArgumentException("Table column width weights must be positive finite values.");
                    }
                } else if (autoFitWeights != null && column < autoFitWeights.Length) {
                    weight = autoFitWeights[column];
                }

                if (autoFitWeights != null && minWidth.HasValue) {
                    AutoFitColumnProfile profile = autoFitProfiles != null && column < autoFitProfiles.Length
                        ? autoFitProfiles[column]
                        : default;
                    weight = ResolveAutoFitFlexibleWeight(weight, minWidth.Value, profile);
                }

                columnWeights[column] = weight;
                totalWeight += weight;
            }

            double tableInnerWidth = tableWidth - (columns - 1) * columnGap;
            if (tableInnerWidth <= 0.001 || double.IsNaN(tableInnerWidth) || double.IsInfinity(tableInnerWidth)) {
                throw new ArgumentException("Table cell spacing must leave a positive table width.");
            }

            fixedWidthTotal = FitFixedTableColumnsToAvailableWidth(columnWidths, fixedColumns, minWidths, fixedWidthTotal, tableInnerWidth);
            if (style.PreserveWidth && totalWeight <= 0D) {
                fixedWidthTotal = ExpandFixedTableColumnsToAvailableWidth(columnWidths, fixedColumns, maxWidths, fixedWidthTotal, tableInnerWidth);
            }
            if (totalWeight <= 0) {
                tableInnerWidth = fixedWidthTotal;
                tableWidth = tableInnerWidth + (columns - 1) * columnGap;
            }

            double remainingWidth = Math.Max(0, tableInnerWidth - fixedWidthTotal);
            DistributeFlexibleColumns(columnWidths, columnWeights, fixedColumns, minWidths, maxWidths, remainingWidth);
            double usedTableInnerWidth = columnWidths.Sum();
            if (!style.PreserveWidth && usedTableInnerWidth < tableInnerWidth - 0.001) {
                tableInnerWidth = usedTableInnerWidth;
                tableWidth = tableInnerWidth + (columns - 1) * columnGap;
            }

            return new PreparedTableColumns(tableWidth, columnWidths);
        }

    }
}
