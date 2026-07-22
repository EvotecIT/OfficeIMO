using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private bool CanUseSequentialRangeFastPath(string operationName, int workload, OfficeIMO.Excel.ExecutionMode? mode) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                return true;
            }

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel) {
                return false;
            }

            if (policy.OnDecision != null) {
                return false;
            }

            int threshold = policy.OperationThresholds.TryGetValue(operationName, out int value)
                ? value
                : policy.ParallelThreshold;
            return workload <= threshold;
        }

        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "Inferred worksheet column types are normalized to OfficeIMO's closed scalar set and are used only as DataColumn conversion tokens.")]
        private void FillDataTableSequential(
            DataTable dt,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            bool headersInFirstRow,
            CancellationToken ct) {
            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            var rowValues = new object?[dataRowCount][];

            if (headersInFirstRow && rows > 0) {
                var headerValues = new object?[cols];
                FillSequentialBuffers(r1, c1, r2, c2, cols, startRow, headerValues, rowValues, ct);
                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                Type[]? columnTypes = _opt.InferDataTableColumnTypes ? InferDataTableColumnTypes(rowValues, cols) : null;
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], columnTypes?[c] ?? typeof(object));
                }
            } else {
                FillSequentialBuffers(r1, c1, r2, c2, cols, startRow, null, rowValues, ct);
                Type[]? columnTypes = _opt.InferDataTableColumnTypes ? InferDataTableColumnTypes(rowValues, cols) : null;
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add($"Column{c + 1}", columnTypes?[c] ?? typeof(object));
                }
            }

            dt.MinimumCapacity = Math.Max(dt.MinimumCapacity, dataRowCount);
            dt.BeginLoadData();
            try {
                AddBufferedRowsToDataTable(dt, rowValues, null, dataRowCount, cols, ct);
            } finally {
                dt.EndLoadData();
            }
        }

        private bool TryFillDataTableSequentialSinglePass(
            DataTable dt,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            bool headersInFirstRow,
            CancellationToken ct) {
            if (_opt.InferDataTableColumnTypes) {
                return false;
            }

            bool canCancel = ct.CanBeCanceled;
            bool columnsReady = false;
            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            DataRow[]? dataRows = null;
            int convertedCells = 0;

            if (!headersInFirstRow) {
                AddDefaultColumns(dt, cols);
                dataRows = CreateDataRows(dt, dataRowCount);
                columnsReady = true;
            }

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (row.RowIndex == null) {
                    return false;
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (headersInFirstRow && rowIndex == r1) {
                    var headerValues = new object?[cols];
                    FillHeaderValues(row, c1, c2, cols, headerValues, ct, ref convertedCells);
                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    for (int c = 0; c < cols; c++) {
                        dt.Columns.Add(headers[c], typeof(object));
                    }

                    dataRows = CreateDataRows(dt, dataRowCount);
                    columnsReady = true;
                    continue;
                }

                if (!columnsReady || dataRows == null) {
                    return false;
                }

                int rr = rowIndex - r1 - startRow;
                if ((uint)rr >= (uint)dataRows.Length) {
                    continue;
                }

                FillDataRow(row, c1, c2, cols, dataRows[rr], ct, ref convertedCells);
            }

            if (dataRows != null) {
                AddRows(dt, dataRows);
            }

            return columnsReady || !headersInFirstRow;

            static void AddDefaultColumns(DataTable table, int columnCount) {
                for (int c = 0; c < columnCount; c++) {
                    table.Columns.Add($"Column{c + 1}", typeof(object));
                }
            }

            static DataRow[] CreateDataRows(DataTable table, int rowCount) {
                var rowsBuffer = new DataRow[rowCount];
                for (int r = 0; r < rowCount; r++) {
                    rowsBuffer[r] = table.NewRow();
                }

                return rowsBuffer;
            }

            static void AddRows(DataTable table, DataRow[] rowsBuffer) {
                for (int r = 0; r < rowsBuffer.Length; r++) {
                    table.Rows.Add(rowsBuffer[r]);
                }
            }

            void FillHeaderValues(Row row, int firstColumn, int lastColumn, int columnCount, object?[] headers, CancellationToken token, ref int visitedCells) {
                bool rowCanCancel = token.CanBeCanceled;
                foreach (var cell in row.Elements<Cell>()) {
                    if (rowCanCancel && (++visitedCells & 1023) == 0) {
                        token.ThrowIfCancellationRequested();
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (columnIndex < firstColumn || columnIndex > lastColumn) {
                        continue;
                    }

                    int cc = columnIndex - firstColumn;
                    if ((uint)cc >= (uint)columnCount) {
                        continue;
                    }

                    if (TryConvertCell(cell, out object? value)) {
                        headers[cc] = value;
                    }
                }
            }

            void FillDataRow(Row row, int firstColumn, int lastColumn, int columnCount, DataRow dataRow, CancellationToken token, ref int visitedCells) {
                bool rowCanCancel = token.CanBeCanceled;
                foreach (var cell in row.Elements<Cell>()) {
                    if (rowCanCancel && (++visitedCells & 1023) == 0) {
                        token.ThrowIfCancellationRequested();
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (columnIndex < firstColumn || columnIndex > lastColumn) {
                        continue;
                    }

                    int cc = columnIndex - firstColumn;
                    if ((uint)cc >= (uint)columnCount) {
                        continue;
                    }

                    if (TryConvertCell(cell, out object? value)) {
                        dataRow[cc] = value ?? DBNull.Value;
                    }
                }
            }
        }

        private List<Dictionary<string, object?>> ReadObjectsSequential(int r1, int c1, int r2, int c2, int rows, int cols, CancellationToken ct) {
            var headerValues = new object?[cols];
            int dataRowCount = rows - 1;
            var rowValues = new object?[dataRowCount][];

            FillSequentialBuffers(r1, c1, r2, c2, cols, startRow: 1, headerValues, rowValues, ct);

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            var result = new List<Dictionary<string, object?>>(dataRowCount);
            for (int r = 0; r < dataRowCount; r++) {
                var values = rowValues[r];
                var dict = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++) {
                    dict.Add(headers[c], values?[c]);
                }

                result.Add(dict);
            }

            return result;
        }
    }
}
