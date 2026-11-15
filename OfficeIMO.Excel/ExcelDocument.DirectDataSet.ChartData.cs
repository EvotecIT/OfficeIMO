using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Reads a chart's cached values straight from a table whose rows still wait for the direct writer, so charting it
        /// does not write every row into the worksheet first. Each value reads as the text its worksheet cell would hold.
        /// Returns false for anything else, and the caller reads the cells.
        /// </summary>
        internal bool TryReadDeferredTabularChartData(ExcelSheet sheet, ExcelChartDataRange range, out ExcelChartData? data) {
            data = null;
            if (range.Orientation != ExcelChartDataOrientation.Vertical
                || !string.Equals(range.SheetName, sheet.Name, StringComparison.Ordinal)) {
                return false;
            }

            if (!TryReadDeferredTabularCellTexts(sheet, range.CategoryCount,
                    index => (range.CategoryStartRow + index, range.CategoryStartColumn), out string?[] categories)) {
                return false;
            }

            var series = new List<ExcelChartSeries>(range.SeriesCount);
            for (int s = 0; s < range.SeriesCount; s++) {
                string name = "Series " + (s + 1).ToString(CultureInfo.InvariantCulture);
                if (range.HasHeaderRow) {
                    if (!TryReadDeferredTabularCellTexts(sheet, 1, _ => (range.StartRow, range.SeriesStartColumn + s), out string?[] header)) {
                        return false;
                    }

                    if (!string.IsNullOrWhiteSpace(header[0])) name = header[0]!;
                }

                int column = range.SeriesStartColumn + s;
                if (!TryReadDeferredTabularCellTexts(sheet, range.CategoryCount, index => (range.CategoryStartRow + index, column), out string?[] raw)) {
                    return false;
                }

                series.Add(new ExcelChartSeries(name, raw.Select(ParseChartNumber)));
            }

            data = new ExcelChartData(categories.Select(text => text ?? string.Empty), series);
            return true;
        }

        /// <summary>The number a chart cache takes from a cell's text, as linked range charts read it.</summary>
        internal static double ParseChartNumber(string? text) =>
            text != null && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) ? value : 0D;

        /// <summary>
        /// Reads cell texts on <paramref name="sheet"/> from a table whose rows still wait for the direct writer: each as the
        /// text its cell would hold, or null for an empty cell. Returns false when the rows are not deferred, a cell lies
        /// outside the table, or a value is anything but plain text or a plain number. With
        /// <paramref name="numbersReadAsNumbers"/>, it also returns false for a number in a date-formatted column, which a
        /// cell value snapshot reads as a date rather than a number.
        /// </summary>
        internal bool TryReadDeferredTabularCellTexts(ExcelSheet sheet, int count, Func<int, (int Row, int Column)> cellAt, out string?[] texts,
            bool numbersReadAsNumbers = false) {
            texts = Array.Empty<string?>();
            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || !candidate.IsDeferred || candidate.Model.Sheets.Count != 1) {
                return false;
            }

            var sheetModel = candidate.Model.Sheets[0];
            if (!string.Equals(sheetModel.SheetName, sheet.Name, StringComparison.Ordinal)
                || sheetModel.StreamingLayout != null
                || !ExcelA1.TryParseRange(sheetModel.Range, out int firstRow, out int firstColumn, out _, out _)) {
                return false;
            }

            var table = sheetModel.Table;
            int dataStartRow = firstRow + (sheetModel.IncludeHeaders ? 1 : 0);
            var result = new string?[count];
            for (int i = 0; i < count; i++) {
                (int row, int column) = cellAt(i);
                int columnIndex = column - firstColumn;
                if (columnIndex < 0 || columnIndex >= table.ColumnCount) return false;
                if (sheetModel.IncludeHeaders && row == firstRow) {
                    result[i] = table.GetColumnName(columnIndex);
                    if (!IsPlainText(result[i])) return false;
                    continue;
                }

                int rowIndex = row - dataStartRow;
                if (rowIndex < 0 || rowIndex >= table.RowCount) return false;
                object? value = table.GetValue(rowIndex, columnIndex);
                if (value == null) continue;
                if (value is string text) {
                    if (!IsPlainText(text)) return false;
                    result[i] = text;
                    continue;
                }

                // Plain numbers read back as their stored text; dates, durations, Booleans and other values keep the cell read.
                if (value is not (double or float or decimal or int or long or short or sbyte or uint or ulong or ushort or byte)) return false;
                var (cellValue, cellType) = CoerceValueHelper.Coerce(value, static shared => new CellValue(shared), dateSystem: DateSystem);
                if (cellType != CellValues.Number) return false;
                if (numbersReadAsNumbers && sheetModel.ColumnNumberFormats is { } formats && columnIndex < formats.Count
                    && ExcelNumberFormatClassifier.LooksLikeDateFormat(formats[columnIndex])) {
                    return false;
                }

                result[i] = cellValue.Text;
            }

            texts = result;
            return true;
        }

        /// <summary>Text a worksheet cell stores unchanged: no control characters or escape-like sequences to sanitize.</summary>
        private static bool IsPlainText(string? text) {
            if (text == null) return true;
            for (int i = 0; i < text.Length; i++) {
                char current = text[i];
                if (current < ' ' && current != '\t' && current != '\n' && current != '\r') return false;
                if (char.IsSurrogate(current) || current == '￾' || current == '￿') return false;
                if (current == '_' && i + 1 < text.Length && (text[i + 1] == 'x' || text[i + 1] == 'X')) return false;
            }

            return true;
        }
    }
}
