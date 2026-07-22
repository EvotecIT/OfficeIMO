using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal sealed class ExcelWorksheetGeometryIndex {
        private const int MaximumColumnDefinitionWork = 100_000;
        private const int MaximumIndexedRows = 100_000;
        private readonly X.Column?[] _columns;
        private readonly IReadOnlyDictionary<int, X.Row> _rows;
        private readonly double _defaultColumnWidth;
        private readonly double _defaultRowHeightPoints;

        private ExcelWorksheetGeometryIndex(
            X.Column?[] columns,
            IReadOnlyDictionary<int, X.Row> rows,
            double defaultColumnWidth,
            double defaultRowHeightPoints) {
            _columns = columns;
            _rows = rows;
            _defaultColumnWidth = defaultColumnWidth;
            _defaultRowHeightPoints = defaultRowHeightPoints;
        }

        internal static ExcelWorksheetGeometryIndex Create(WorksheetPart? worksheetPart) {
            X.Worksheet? worksheet = worksheetPart?.Worksheet;
            X.SheetFormatProperties? format = worksheet?.GetFirstChild<X.SheetFormatProperties>();
            double defaultColumnWidth = format?.DefaultColumnWidth?.Value > 0
                ? format.DefaultColumnWidth.Value
                : 8.43D;
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

            var rows = new Dictionary<int, X.Row>();
            foreach (X.Row row in worksheet?
                .GetFirstChild<X.SheetData>()?
                .Elements<X.Row>() ?? Enumerable.Empty<X.Row>()) {
                if (rows.Count >= MaximumIndexedRows) {
                    break;
                }

                if (row.RowIndex?.Value is uint rowIndex && rowIndex > 0U && rowIndex <= 1048576U) {
                    int index = (int)rowIndex;
                    if (!rows.ContainsKey(index)) {
                        rows[index] = row;
                    }
                }
            }

            return new ExcelWorksheetGeometryIndex(columns, rows, defaultColumnWidth, defaultRowHeight);
        }

        internal int GetSimpleColumnWidthPixels(int columnIndex) {
            X.Column? column = columnIndex > 0 && columnIndex < _columns.Length ? _columns[columnIndex] : null;
            if (column?.Hidden?.Value == true) {
                return 0;
            }

            double width = column?.Width?.Value > 0 && column.CustomWidth?.Value == true
                ? column.Width.Value
                : _defaultColumnWidth;
            return Math.Max(1, (int)Math.Round(Math.Round((width * 7D) + 5D, 2)));
        }

        internal int GetColumnWidthPixels(int columnIndex, double maximumDigitWidth) {
            X.Column? column = columnIndex > 0 && columnIndex < _columns.Length ? _columns[columnIndex] : null;
            if (column?.Hidden?.Value == true) {
                return 0;
            }

            double width = column?.Width?.Value > 0 && column.CustomWidth?.Value == true
                ? column.Width.Value
                : _defaultColumnWidth;
            double pixels = Math.Truncate((256D * width + Math.Truncate(128D / maximumDigitWidth)) / 256D * maximumDigitWidth);
            return Math.Max(1, (int)Math.Round(pixels));
        }

        internal int GetRowHeightPixels(int rowIndex) {
            _rows.TryGetValue(rowIndex, out X.Row? row);
            if (row?.Hidden?.Value == true) {
                return 0;
            }

            double heightPoints = row?.Height?.Value > 0 && row.CustomHeight?.Value == true
                ? row.Height.Value
                : _defaultRowHeightPoints;
            return Math.Max(1, (int)Math.Round(heightPoints * 96D / 72D));
        }
    }
}
