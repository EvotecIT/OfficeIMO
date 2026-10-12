using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal sealed class ExcelWorksheetGeometryIndex {
        private const int MaximumColumnDefinitionWork = 100_000;
        private const int MaximumIndexedRows = 100_000;
        private const int MaximumOverflowRowDefinitionWork = 1_048_576;
        private readonly X.Column?[] _columns;
        private readonly IReadOnlyDictionary<int, X.Row> _rows;
        private readonly X.SheetData? _sheetData;
        private readonly bool _rowIndexMayBeTruncated;
        private readonly X.SheetFormatProperties? _sheetFormat;
        private readonly double _maximumDigitWidth;
        private readonly double _defaultRowHeightPoints;
        private double[]? _overflowRowHeights;

        private ExcelWorksheetGeometryIndex(
            X.Column?[] columns,
            IReadOnlyDictionary<int, X.Row> rows,
            X.SheetData? sheetData,
            bool rowIndexMayBeTruncated,
            X.SheetFormatProperties? sheetFormat,
            double maximumDigitWidth,
            double defaultRowHeightPoints) {
            _columns = columns;
            _rows = rows;
            _sheetData = sheetData;
            _rowIndexMayBeTruncated = rowIndexMayBeTruncated;
            _sheetFormat = (X.SheetFormatProperties?)sheetFormat?.CloneNode(true);
            _maximumDigitWidth = maximumDigitWidth;
            _defaultRowHeightPoints = defaultRowHeightPoints;
        }

        internal static ExcelWorksheetGeometryIndex Create(WorksheetPart? worksheetPart) {
            X.Worksheet? worksheet = worksheetPart?.Worksheet;
            X.SheetFormatProperties? format = worksheet?.GetFirstChild<X.SheetFormatProperties>();
            double defaultRowHeight = format?.DefaultRowHeight?.Value > 0
                ? format.DefaultRowHeight.Value
                : 15D;

            var columns = new X.Column?[16385];
            int remainingColumnWork = MaximumColumnDefinitionWork;
            foreach (X.Column definition in worksheet?
                .GetFirstChild<X.Columns>()?
                .Elements<X.Column>() ?? Enumerable.Empty<X.Column>()) {
                if (remainingColumnWork <= 0) {
                    break;
                }
                remainingColumnWork--;

                if (definition.Min?.Value is not uint minimum ||
                    definition.Max?.Value is not uint maximum ||
                    minimum == 0U ||
                    minimum > 16384U ||
                    maximum < minimum) {
                    continue;
                }

                int start = (int)minimum;
                int end = (int)Math.Min(16384U, maximum);

                for (int column = start; column <= end && remainingColumnWork > 0; column++, remainingColumnWork--) {
                    columns[column] ??= definition;
                }
            }

            X.SheetData? sheetData = worksheet?.GetFirstChild<X.SheetData>();
            var rows = new Dictionary<int, X.Row>();
            int processedRowDefinitions = 0;
            bool rowIndexMayBeTruncated = false;
            foreach (X.Row row in sheetData?.Elements<X.Row>() ?? Enumerable.Empty<X.Row>()) {
                if (processedRowDefinitions >= MaximumIndexedRows) {
                    rowIndexMayBeTruncated = true;
                    break;
                }
                processedRowDefinitions++;

                if (row.RowIndex?.Value is uint rowIndex && rowIndex > 0U && rowIndex <= 1048576U) {
                    int index = (int)rowIndex;
                    if (!rows.ContainsKey(index)) {
                        rows[index] = row;
                    }
                }
            }

            return new ExcelWorksheetGeometryIndex(
                columns,
                rows,
                sheetData,
                rowIndexMayBeTruncated,
                format,
                ExcelColumnGeometry.GetMaximumDigitWidth(worksheetPart),
                defaultRowHeight);
        }

        internal int GetSimpleColumnWidthPixels(int columnIndex) =>
            GetColumnWidthPixels(columnIndex, _maximumDigitWidth);

        internal int GetColumnWidthPixels(int columnIndex, double maximumDigitWidth) {
            X.Column? column = columnIndex > 0 && columnIndex < _columns.Length ? _columns[columnIndex] : null;
            return ExcelColumnGeometry.GetWidthPixels(column, _sheetFormat, maximumDigitWidth);
        }

        /// <summary>Uses authored worksheet geometry, retaining the export fallback for unspecified widths.</summary>
        internal double GetExportColumnWidthPixels(int columnIndex, double fallbackPixels) {
            X.Column? column = columnIndex > 0 && columnIndex < _columns.Length ? _columns[columnIndex] : null;
            if (!(column?.Width?.Value > 0D && column.CustomWidth?.Value == true) &&
                !(_sheetFormat?.DefaultColumnWidth?.Value > 0D) && !(_sheetFormat?.BaseColumnWidth?.Value > 0U))
                return fallbackPixels;
            return ExcelColumnGeometry.GetWidthPixels(column, _sheetFormat, _maximumDigitWidth, includeHidden: true);
        }

        internal int GetRowHeightPixels(int rowIndex) {
            _rows.TryGetValue(rowIndex, out X.Row? row);
            if (row?.Hidden?.Value == true) {
                return 0;
            }

            double heightPoints = _defaultRowHeightPoints;
            if (row != null) {
                if (row.Height?.Value > 0 && row.CustomHeight?.Value == true) {
                    heightPoints = row.Height.Value;
                }
            } else {
                double overflowHeight = GetOverflowRowHeight(rowIndex);
                if (overflowHeight < 0D) {
                    return 0;
                }

                if (overflowHeight > 0D) {
                    heightPoints = overflowHeight;
                }
            }

            return Math.Max(1, (int)Math.Round(heightPoints * 96D / 72D));
        }

        private double GetOverflowRowHeight(int rowIndex) {
            if (!_rowIndexMayBeTruncated || rowIndex <= 0 || rowIndex > 1048576) {
                return 0D;
            }

            EnsureOverflowRowsIndexed();
            return _overflowRowHeights![rowIndex];
        }

        private void EnsureOverflowRowsIndexed() {
            if (_overflowRowHeights != null) {
                return;
            }

            var heights = new double[1048577];
            var assigned = new bool[1048577];
            int processedRowDefinitions = 0;
            foreach (X.Row row in _sheetData?.Elements<X.Row>() ?? Enumerable.Empty<X.Row>()) {
                if (processedRowDefinitions >= MaximumOverflowRowDefinitionWork) {
                    break;
                }
                processedRowDefinitions++;

                if (row.RowIndex?.Value is not uint rowIndex || rowIndex == 0U || rowIndex > 1048576U ||
                    _rows.ContainsKey((int)rowIndex) || assigned[rowIndex]) {
                    continue;
                }

                assigned[rowIndex] = true;
                heights[rowIndex] = row.Hidden?.Value == true
                    ? -1D
                    : row.Height?.Value > 0D && row.CustomHeight?.Value == true
                        ? row.Height.Value
                        : 0D;
            }

            _overflowRowHeights = heights;
        }
    }
}
