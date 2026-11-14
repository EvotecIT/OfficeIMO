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
            var candidate = _directDataSetSaveCandidate;
            if (candidate == null || !candidate.IsValid || !candidate.IsDeferred || candidate.Model.Sheets.Count != 1
                || range.Orientation != ExcelChartDataOrientation.Vertical
                || !string.Equals(range.SheetName, sheet.Name, StringComparison.Ordinal)) {
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
            bool TryGetText(int row, int column, out string? text) {
                text = null;
                int columnIndex = column - firstColumn;
                if (columnIndex < 0 || columnIndex >= table.ColumnCount) return false;
                if (sheetModel.IncludeHeaders && row == firstRow) {
                    text = table.GetColumnName(columnIndex);
                    return IsPlainText(text);
                }

                int rowIndex = row - dataStartRow;
                if (rowIndex < 0 || rowIndex >= table.RowCount) return false;
                object? value = table.GetValue(rowIndex, columnIndex);
                if (value == null) return true;
                if (value is string stringValue) {
                    text = stringValue;
                    return IsPlainText(text);
                }

                // Plain numbers read back as their stored text; dates, durations, Booleans and other values keep the cell read.
                if (value is not (double or float or decimal or int or long or short or sbyte or uint or ulong or ushort or byte)) return false;
                var (cellValue, cellType) = CoerceValueHelper.Coerce(value, static shared => new CellValue(shared), dateSystem: DateSystem);
                if (cellType != CellValues.Number) return false;
                text = cellValue.Text;
                return true;
            }

            var categories = new List<string>(range.CategoryCount);
            for (int i = 0; i < range.CategoryCount; i++) {
                if (!TryGetText(range.CategoryStartRow + i, range.CategoryStartColumn, out string? category)) return false;
                categories.Add(category ?? string.Empty);
            }

            var series = new List<ExcelChartSeries>(range.SeriesCount);
            for (int s = 0; s < range.SeriesCount; s++) {
                string name = "Series " + (s + 1).ToString(CultureInfo.InvariantCulture);
                if (range.HasHeaderRow) {
                    if (!TryGetText(range.StartRow, range.SeriesStartColumn + s, out string? header)) return false;
                    if (!string.IsNullOrWhiteSpace(header)) name = header!;
                }

                var values = new List<double>(range.CategoryCount);
                for (int i = 0; i < range.CategoryCount; i++) {
                    if (!TryGetText(range.CategoryStartRow + i, range.SeriesStartColumn + s, out string? raw)) return false;
                    values.Add(raw != null && double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double value) ? value : 0D);
                }

                series.Add(new ExcelChartSeries(name, values));
            }

            data = new ExcelChartData(categories, series);
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
