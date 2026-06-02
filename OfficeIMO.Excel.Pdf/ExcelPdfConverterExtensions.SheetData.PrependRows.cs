using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Excel.Pdf {
    public static partial class ExcelPdfConverterExtensions {
        private static object?[,] PrependRows(object?[,] topRows, object?[,] bodyRows) {
            int topRowCount = topRows.GetLength(0);
            int bodyRowCount = bodyRows.GetLength(0);
            int columnCount = bodyRows.GetLength(1);
            var result = new object?[topRowCount + bodyRowCount, columnCount];
            CopyRows(topRows, result, 0, columnCount);
            CopyRows(bodyRows, result, topRowCount, columnCount);
            return result;
        }

        private static ExcelCellStyleSnapshot?[,]? PrependRows(ExcelCellStyleSnapshot?[,]? topRows, ExcelCellStyleSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelCellStyleSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static ExcelHyperlinkSnapshot?[,]? PrependRows(ExcelHyperlinkSnapshot?[,]? topRows, ExcelHyperlinkSnapshot?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new ExcelHyperlinkSnapshot?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static string?[,]? PrependRows(string?[,]? topRows, string?[,]? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new string?[topRowCount + bodyRowCount, columnCount];
            if (topRows != null) {
                CopyRows(topRows, result, 0, columnCount);
            }

            if (bodyRows != null) {
                CopyRows(bodyRows, result, topRowCount, columnCount);
            }

            return result;
        }

        private static MergeLayoutData? PrependRows(MergeLayoutData? topRows, MergeLayoutData? bodyRows, int topRowCount, int bodyRowCount, int columnCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var result = new MergeLayoutData(topRowCount + bodyRowCount, columnCount);
            topRows?.CopyTo(result, 0);
            bodyRows?.CopyTo(result, topRowCount);
            return result.HasAny ? result : null;
        }

        private static RowLayoutData? PrependRows(RowLayoutData? topRows, RowLayoutData? bodyRows, int topRowCount, int bodyRowCount) {
            if (topRows == null && bodyRows == null) {
                return null;
            }

            var minHeights = new List<double?>(topRowCount + bodyRowCount);
            if (topRows != null) {
                minHeights.AddRange(topRows.MinHeights);
            } else {
                for (int row = 0; row < topRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            if (bodyRows != null) {
                minHeights.AddRange(bodyRows.MinHeights);
            } else {
                for (int row = 0; row < bodyRowCount; row++) {
                    minHeights.Add(null);
                }
            }

            return minHeights.Any(height => height.HasValue) ? new RowLayoutData(minHeights) : null;
        }

        private static void CopyRows(object?[,] source, object?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelCellStyleSnapshot?[,] source, ExcelCellStyleSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(ExcelHyperlinkSnapshot?[,] source, ExcelHyperlinkSnapshot?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static void CopyRows(string?[,] source, string?[,] target, int targetRowOffset, int columnCount) {
            int rowCount = source.GetLength(0);
            int sourceColumnCount = source.GetLength(1);
            for (int row = 0; row < rowCount; row++) {
                for (int column = 0; column < columnCount; column++) {
                    target[targetRowOffset + row, column] = column < sourceColumnCount ? source[row, column] : null;
                }
            }
        }

        private static string NormalizeA1Range(string range) {
            string withoutSheet = StripSheetPrefix(range).Replace("$", string.Empty);
            if (!A1.TryParseRange(withoutSheet, out int r1, out int c1, out int r2, out int c2)) {
                (int Row, int Col) cell = A1.ParseCellRef(withoutSheet);
                if (cell.Row <= 0 || cell.Col <= 0) {
                    throw new ArgumentException("Excel PDF export range must be a valid A1 range.", nameof(range));
                }

                r1 = r2 = cell.Row;
                c1 = c2 = cell.Col;
            }

            return ToA1Range(r1, c1, r2, c2);
        }

        private static string StripSheetPrefix(string range) {
            int separator = range.LastIndexOf('!');
            return separator >= 0 ? range.Substring(separator + 1) : range;
        }

        private static string ToA1Range(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            string start = A1.ColumnIndexToLetters(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
            string end = A1.ColumnIndexToLetters(lastColumn) + lastRow.ToString(CultureInfo.InvariantCulture);
            return start + ":" + end;
        }

    }
}
