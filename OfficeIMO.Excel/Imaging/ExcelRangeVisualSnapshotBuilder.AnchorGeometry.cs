using OfficeIMO.Excel.Utilities;
using System.Globalization;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeVisualSnapshotBuilder {
        private static bool TryGetImageAnchor(ExcelWorksheetGeometryIndex geometry, 
            ExcelImageExportOptions options,
            int firstRow,
            int firstColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            int columnIndex,
            int rowIndex,
            int offsetXPixels,
            int offsetYPixels,
            int widthPixels,
            int heightPixels,
            int? toColumnIndex,
            int? toRowIndex,
            int toOffsetXPixels,
            int toOffsetYPixels,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (rowIndex <= 0 || columnIndex <= 0 || columnsByIndex.Count == 0 || rowsByIndex.Count == 0) {
                return false;
            }

            int lastVisibleColumn = columnsByIndex.Keys.Max();
            int lastVisibleRow = rowsByIndex.Keys.Max();
            if (DistanceOutsideRange(columnIndex, firstColumn, lastVisibleColumn) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                DistanceOutsideRange(rowIndex, firstRow, lastVisibleRow) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                (toColumnIndex.HasValue &&
                    (DistanceOutsideRange(toColumnIndex.Value, firstColumn, lastVisibleColumn) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                     Math.Abs((long)toColumnIndex.Value - columnIndex) > ExcelImageExportLimits.MaximumAnchorSpanCells)) ||
                (toRowIndex.HasValue &&
                    (DistanceOutsideRange(toRowIndex.Value, firstRow, lastVisibleRow) > ExcelImageExportLimits.MaximumAnchorSpanCells ||
                     Math.Abs((long)toRowIndex.Value - rowIndex) > ExcelImageExportLimits.MaximumAnchorSpanCells))) {
                return false;
            }

            offsetXPixels = ExcelImageExportLimits.ClampOffsetPixels(offsetXPixels);
            offsetYPixels = ExcelImageExportLimits.ClampOffsetPixels(offsetYPixels);
            toOffsetXPixels = ExcelImageExportLimits.ClampOffsetPixels(toOffsetXPixels);
            toOffsetYPixels = ExcelImageExportLimits.ClampOffsetPixels(toOffsetYPixels);
            widthPixels = ExcelImageExportLimits.ClampExtentPixels(widthPixels);
            heightPixels = ExcelImageExportLimits.ClampExtentPixels(heightPixels);
            x = ResolveRelativeColumnOffset(geometry, firstColumn, columnIndex, offsetXPixels, columnDefinitions, options);
            y = ResolveRelativeRowOffset(firstRow, rowIndex, offsetYPixels, rowDefinitions, defaultRowsHidden, options);
            width = ResolveImageWidth(geometry, options, firstColumn, columnDefinitions, columnIndex, offsetXPixels, widthPixels, toColumnIndex, toOffsetXPixels);
            height = ResolveImageHeight(options, firstRow, rowDefinitions, defaultRowsHidden, rowIndex, offsetYPixels, heightPixels, toRowIndex, toOffsetYPixels);

            double rangeWidth = ResolveVisualWidth(columnsByIndex);
            double rangeHeight = ResolveVisualHeight(rowsByIndex);
            return x < rangeWidth &&
                x + width > 0D &&
                y < rangeHeight &&
                y + height > 0D;
        }

        private static long DistanceOutsideRange(int value, int first, int last) {
            if (value < first) {
                return (long)first - value;
            }

            return value > last ? (long)value - last : 0L;
        }

        private static bool TryGetAnchor(
            IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex,
            IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex,
            int columnIndex,
            int rowIndex,
            int widthPixels,
            int heightPixels,
            out double x,
            out double y,
            out double width,
            out double height) {
            x = 0D;
            y = 0D;
            width = 0D;
            height = 0D;
            if (!columnsByIndex.TryGetValue(columnIndex, out ExcelVisualColumn? column) ||
                !rowsByIndex.TryGetValue(rowIndex, out ExcelVisualRow? row)) {
                return false;
            }

            x = column.X;
            y = row.Y;
            width = Math.Max(1D, widthPixels);
            height = Math.Max(1D, heightPixels);
            return true;
        }

        private static double ResolveRelativeColumnOffset(ExcelWorksheetGeometryIndex geometry, 
            int firstColumn,
            int anchorColumn,
            int offsetPixels,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            ExcelImageExportOptions options) {
            double x = 0D;
            if (anchorColumn >= firstColumn) {
                for (int column = firstColumn; column < anchorColumn; column++) {
                    x += ResolveVisibleColumnWidth(geometry, column, columnDefinitions, options);
                }
            } else {
                for (int column = anchorColumn; column < firstColumn; column++) {
                    x -= ResolveVisibleColumnWidth(geometry, column, columnDefinitions, options);
                }
            }

            return x + offsetPixels;
        }

        private static double ResolveImageWidth(ExcelWorksheetGeometryIndex geometry, 
            ExcelImageExportOptions options,
            int firstColumn,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            int anchorColumn,
            int offsetPixels,
            int widthPixels,
            int? toColumnIndex,
            int toOffsetPixels) {
            if (toColumnIndex.HasValue) {
                double from = ResolveRelativeColumnOffset(geometry, firstColumn, anchorColumn, offsetPixels, columnDefinitions, options);
                double to = ResolveRelativeColumnOffset(geometry, firstColumn, toColumnIndex.Value, toOffsetPixels, columnDefinitions, options);
                if (to > from) {
                    return Math.Min(ExcelImageExportLimits.MaximumAnchorExtentPixels, Math.Max(1D, to - from));
                }
            }

            return Math.Max(1D, widthPixels);
        }

        private static double ResolveRelativeRowOffset(
            int firstRow,
            int anchorRow,
            int offsetPixels,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            ExcelImageExportOptions options) {
            double y = 0D;
            if (anchorRow >= firstRow) {
                for (int row = firstRow; row < anchorRow; row++) {
                    y += ResolveVisibleRowHeight(row, rowDefinitions, defaultRowsHidden, options);
                }
            } else {
                for (int row = anchorRow; row < firstRow; row++) {
                    y -= ResolveVisibleRowHeight(row, rowDefinitions, defaultRowsHidden, options);
                }
            }

            return y + offsetPixels;
        }

        private static double ResolveImageHeight(
            ExcelImageExportOptions options,
            int firstRow,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            int anchorRow,
            int offsetPixels,
            int heightPixels,
            int? toRowIndex,
            int toOffsetPixels) {
            if (toRowIndex.HasValue) {
                double from = ResolveRelativeRowOffset(firstRow, anchorRow, offsetPixels, rowDefinitions, defaultRowsHidden, options);
                double to = ResolveRelativeRowOffset(firstRow, toRowIndex.Value, toOffsetPixels, rowDefinitions, defaultRowsHidden, options);
                if (to > from) {
                    return Math.Min(ExcelImageExportLimits.MaximumAnchorExtentPixels, Math.Max(1D, to - from));
                }
            }

            return Math.Max(1D, heightPixels);
        }

        private static double ResolveVisibleColumnWidth(ExcelWorksheetGeometryIndex geometry, int column, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions, ExcelImageExportOptions options) {
            ExcelColumnSnapshot? definition = columnDefinitions.FirstOrDefault(item => column >= item.StartIndex && column <= item.EndIndex);
            if (definition?.Hidden == true && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveColumnWidth(geometry, definition, options);
        }

        private static double ResolveVisibleRowHeight(int row, IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions, bool defaultRowsHidden, ExcelImageExportOptions options) {
            rowDefinitions.TryGetValue(row, out ExcelRowSnapshot? definition);
            if (IsHiddenRow(row, rowDefinitions, defaultRowsHidden) && !options.IncludeHidden) {
                return 0D;
            }

            return ResolveRowHeight(definition, options);
        }

        private static double ResolveVisualWidth(IReadOnlyDictionary<int, ExcelVisualColumn> columnsByIndex) =>
            columnsByIndex.Count == 0 ? 0D : columnsByIndex.Values.Max(column => column.X + column.Width);

        private static double ResolveVisualHeight(IReadOnlyDictionary<int, ExcelVisualRow> rowsByIndex) =>
            rowsByIndex.Count == 0 ? 0D : rowsByIndex.Values.Max(row => row.Y + row.Height);

        private static void ReportHiddenRowColumnOmissions(
            ExcelSheet sheet,
            string range,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions,
            List<OfficeImageExportDiagnostic> diagnostics) {
            int hiddenRows = 0;
            for (int row = firstRow; row <= lastRow; row++) {
                if (IsHiddenRow(row, rowDefinitions, defaultRowsHidden)) {
                    hiddenRows++;
                }
            }

            int hiddenColumns = CountHiddenColumns(firstColumn, lastColumn, columnDefinitions);
            string source = sheet.Name + "!" + range;
            if (hiddenRows > 0) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HiddenRowsOmitted,
                    hiddenRows.ToString(CultureInfo.InvariantCulture) + " hidden row(s) were omitted from the Excel image export.",
                    source));
            }

            if (hiddenColumns > 0) {
                diagnostics.Add(ExcelImageExportDiagnosticClassifier.Create(
                    OfficeImageExportDiagnosticSeverity.Warning,
                    ExcelImageExportDiagnosticCodes.HiddenColumnsOmitted,
                    hiddenColumns.ToString(CultureInfo.InvariantCulture) + " hidden column(s) were omitted from the Excel image export.",
                    source));
            }
        }

        private static int CountHiddenColumns(int firstColumn, int lastColumn, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            var hiddenColumns = new HashSet<int>();
            foreach (ExcelColumnSnapshot definition in columnDefinitions) {
                if (!definition.Hidden) {
                    continue;
                }

                int start = Math.Max(firstColumn, definition.StartIndex);
                int end = Math.Min(lastColumn, definition.EndIndex);
                for (int column = start; column <= end; column++) {
                    hiddenColumns.Add(column);
                }
            }

            return hiddenColumns.Count;
        }

        private static bool IsHiddenAnchorInRange(
            int rowIndex,
            int columnIndex,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions,
            bool defaultRowsHidden,
            IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            return rowIndex >= firstRow &&
                rowIndex <= lastRow &&
                columnIndex >= firstColumn &&
                columnIndex <= lastColumn &&
                (IsHiddenRow(rowIndex, rowDefinitions, defaultRowsHidden) || IsHiddenColumn(columnIndex, columnDefinitions));
        }

        private static bool IsHiddenRow(int rowIndex, IReadOnlyDictionary<int, ExcelRowSnapshot> rowDefinitions, bool defaultRowsHidden) {
            if (rowDefinitions.TryGetValue(rowIndex, out ExcelRowSnapshot? definition)) {
                return definition.Hidden;
            }

            return defaultRowsHidden;
        }

        private static bool IsHiddenColumn(int columnIndex, IReadOnlyList<ExcelColumnSnapshot> columnDefinitions) {
            return columnDefinitions.Any(definition =>
                definition.Hidden &&
                columnIndex >= definition.StartIndex &&
                columnIndex <= definition.EndIndex);
        }

        private static bool Intersects(int r1, int c1, int r2, int c2, int otherR1, int otherC1, int otherR2, int otherC2) =>
            r1 <= otherR2 && r2 >= otherR1 && c1 <= otherC2 && c2 >= otherC1;

        private static double ResolveColumnWidth(ExcelWorksheetGeometryIndex geometry, ExcelColumnSnapshot? definition, ExcelImageExportOptions options) {
            return geometry.GetExportColumnWidthPixels(definition?.StartIndex ?? 0, options.DefaultColumnWidthPixels);
        }

        private static double ResolveRowHeight(ExcelRowSnapshot? definition, ExcelImageExportOptions options) {
            if (definition?.Height == null) {
                return options.DefaultRowHeightPixels;
            }

            return Math.Max(1D, Math.Round(definition.Height.Value * 96D / 72D, 2));
        }

    }
}
