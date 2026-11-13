using System.Data;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Reads the worksheet used range as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable ToDataTable(bool headersInFirstRow = true, ExcelReaderOptions? options = null, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var reader = _excelDocument.CreateTabularReader(this, effectiveOptions.WithCancellationToken(token));
            var sheetReader = reader.GetSheet(Name);
            string a1Range = sheetReader.GetUsedRangeA1(token);
            return sheetReader.ReadRangeAsDataTable(a1Range, headersInFirstRow, mode, token);
        }

        /// <summary>
        /// Reads an A1 range as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable ToDataTable(string a1Range, bool headersInFirstRow = true, ExcelReaderOptions? options = null, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentNullException(nameof(a1Range));
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var reader = _excelDocument.CreateTabularReader(this, effectiveOptions.WithCancellationToken(token));
            return reader.GetSheet(Name).ReadRangeAsDataTable(a1Range, headersInFirstRow, mode, token);
        }

        /// <summary>
        /// Reads an Excel table as a <see cref="DataTable"/>.
        /// </summary>
        public DataTable TableToDataTable(string tableName, bool? headersInFirstRow = null, ExcelReaderOptions? options = null, ExcelExecutionMode? mode = null, CancellationToken ct = default) {
            if (string.IsNullOrWhiteSpace(tableName)) throw new ArgumentNullException(nameof(tableName));
            ExcelReaderOptions effectiveOptions = options ?? new ExcelReaderOptions();
            using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                ct,
                effectiveOptions.CancellationToken);
            CancellationToken token = linkedCancellation.Token;
            using var reader = _excelDocument.CreateTabularReader(null, effectiveOptions.WithCancellationToken(token));
            return reader.ReadTableAsDataTable(tableName, headersInFirstRow, mode, token);
        }

        private static string BuildInsertedRange(DataTable table, int startRow, int startColumn, bool includeHeaders) {
            int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
            if (table.Columns.Count == 0 || rowCount == 0) {
                return string.Empty;
            }

            return ExcelA1.CellReference(startRow, startColumn) + ":" + ExcelA1.CellReference(startRow + rowCount - 1, startColumn + table.Columns.Count - 1);
        }
    }
}
