using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private const int DenseSnapshotCapacityLimit = 100_000;

        /// <summary>
        /// Returns the used range of the worksheet as an A1 string (e.g., "A1:C10").
        /// If the sheet is empty, returns "A1:A1".
        /// </summary>
        public string GetUsedRangeA1() {
            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            return reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
        }
        /// <summary>
        /// Reads a rectangular A1 range (e.g., "A1:C10") into a dense 2D array of typed values.
        /// </summary>
        public object?[,] ReadRange(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            var height = r2 - r1 + 1;
            var width = c2 - c1 + 1;
            var result = new object?[height, width];

            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = height * width;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide("ReadRange", workload);
            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                FillRangeSequential(result, r1, c1, r2, c2, ct);
                return result;
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRange", decided, ct, workload);

            foreach (var cell in raw) {
                var rr = cell.Row - r1;
                var cc = cell.Col - c1;
                if ((uint)rr < (uint)height && (uint)cc < (uint)width)
                    result[rr, cc] = cell.TypedValue;
            }

            return result;
        }

        /// <summary>
        /// Reads a rectangular range to a DataTable. If headersInFirstRow = true, first row becomes column names.
        /// </summary>
        public DataTable ReadRangeAsDataTable(string a1Range, bool headersInFirstRow = true, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            var dt = new DataTable(_sheetName);
            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (CanUseSequentialRangeFastPath("ReadRangeAsDataTable", rows * cols, mode)) {
                if (TryFillDataTableSequentialSinglePass(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                    return dt;
                }

                FillDataTableSequential(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct);
                return dt;
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRangeAsDataTable", mode, ct, rows * cols);

            if (headersInFirstRow && rows > 0) {
                var headerValues = new object?[cols];
                foreach (var cell in raw) {
                    if (cell.Row != r1) continue;
                    int cc = cell.Col - c1;
                    if ((uint)cc < (uint)cols) {
                        headerValues[cc] = cell.TypedValue;
                    }
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) dt.Columns.Add($"Column{c + 1}", typeof(object));
            }

            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            var dataRows = new DataRow[dataRowCount];
            for (int r = 0; r < dataRowCount; r++) {
                dataRows[r] = dt.NewRow();
                dt.Rows.Add(dataRows[r]);
            }

            foreach (var cell in raw) {
                int rr = cell.Row - r1 - startRow;
                int cc = cell.Col - c1;
                if ((uint)rr < (uint)dataRowCount && (uint)cc < (uint)cols) {
                    dataRows[rr][cc] = cell.TypedValue ?? DBNull.Value;
                }
            }

            return dt;
        }

        /// <summary>
        /// Reads a rectangular range into a sequence of dictionaries using the first row as headers.
        /// </summary>
        public IEnumerable<Dictionary<string, object?>> ReadObjects(string a1Range, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            if (rows <= 1 || cols == 0) {
                return Array.Empty<Dictionary<string, object?>>();
            }

            if (CanUseSequentialRangeFastPath("ReadObjects", rows * cols, mode)) {
                if (TryReadObjectsSequentialSinglePass(r1, c1, r2, c2, rows, cols, ct, out var fastResult)) {
                    return fastResult;
                }

                return ReadObjectsSequential(r1, c1, r2, c2, rows, cols, ct);
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadObjects", mode, ct, rows * cols);

            var headerValues = new object?[cols];
            foreach (var cell in raw) {
                if (cell.Row != r1) continue;
                int cc = cell.Col - c1;
                if ((uint)cc < (uint)cols) {
                    headerValues[cc] = cell.TypedValue;
                }
            }

            var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
            int dataRowCount = rows - 1;
            var rowValues = new object?[dataRowCount][];

            foreach (var cell in raw) {
                if (cell.Row <= r1) continue;
                int rr = cell.Row - r1 - 1;
                int cc = cell.Col - c1;
                if ((uint)rr < (uint)dataRowCount && (uint)cc < (uint)cols) {
                    var values = rowValues[rr];
                    if (values == null) {
                        values = new object?[cols];
                        rowValues[rr] = values;
                    }

                    values[cc] = cell.TypedValue;
                }
            }

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
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add($"Column{c + 1}", typeof(object));
                }

                FillSequentialBuffers(r1, c1, r2, c2, cols, startRow, null, rowValues, ct);
            }

            for (int r = 0; r < dataRowCount; r++) {
                var source = rowValues[r];
                DataRow row = dt.NewRow();
                for (int c = 0; c < cols; c++) {
                    row[c] = source?[c] ?? DBNull.Value;
                }

                dt.Rows.Add(row);
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

        private bool TryReadObjectsSequentialSinglePass(
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<Dictionary<string, object?>> result) {
            int dataRowCount = rows - 1;
            result = new List<Dictionary<string, object?>>(dataRowCount);

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            string[]? headers = null;
            int convertedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) {
                    continue;
                }

                if (rowIndex > r2) {
                    continue;
                }

                if (rowIndex == r1) {
                    var headerValues = new object?[cols];
                    foreach (var cell in row.Elements<Cell>()) {
                        if (canCancel && (++convertedCells & 1023) == 0) {
                            ct.ThrowIfCancellationRequested();
                        }

                        int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                        if (columnIndex < c1 || columnIndex > c2) {
                            continue;
                        }

                        int cc = columnIndex - c1;
                        if ((uint)cc >= (uint)cols) {
                            continue;
                        }

                        if (TryConvertCell(cell, out object? value)) {
                            headerValues[cc] = value;
                        }
                    }

                    headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    for (int r = 0; r < dataRowCount; r++) {
                        if (canCancel && (r & 1023) == 0) {
                            ct.ThrowIfCancellationRequested();
                        }

                        result.Add(CreateEmptyRow(headers));
                    }

                    continue;
                }

                if (headers == null) {
                    return false;
                }

                int rr = rowIndex - r1 - 1;
                if ((uint)rr >= (uint)result.Count) {
                    continue;
                }

                var dict = result[rr];
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++convertedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int columnIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (columnIndex < c1 || columnIndex > c2) {
                        continue;
                    }

                    int cc = columnIndex - c1;
                    if ((uint)cc >= (uint)cols) {
                        continue;
                    }

                    if (TryConvertCell(cell, out object? value)) {
                        dict[headers[cc]] = value;
                    }
                }
            }

            return headers != null;

            Dictionary<string, object?> CreateEmptyRow(string[] rowHeaders) {
                var dict = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
                for (int c = 0; c < cols; c++) {
                    dict.Add(rowHeaders[c], null);
                }

                return dict;
            }
        }

        private void FillSequentialBuffers(
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            int startRow,
            object?[]? headerValues,
            object?[][] rowValues,
            CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < r1) continue;
                if (rowIndex > r2) continue;

                int rr = rowIndex - r1 - startRow;
                bool isHeaderRow = headerValues != null && rowIndex == r1;
                if (!isHeaderRow && (uint)rr >= (uint)rowValues.Length) {
                    continue;
                }

                object?[]? values = null;
                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    int cc = cIndex - c1;
                    if ((uint)cc >= (uint)cols) continue;

                    if (TryConvertCell(cell, out object? value)) {
                        if (isHeaderRow) {
                            headerValues![cc] = value;
                        } else {
                            values ??= rowValues[rr] ??= new object?[cols];
                            values[cc] = value;
                        }
                    }
                }
            }
        }

        private List<CellRaw> SnapshotAndConvertRangeCells(
            int r1,
            int c1,
            int r2,
            int c2,
            string operationName,
            OfficeIMO.Excel.ExecutionMode? mode,
            CancellationToken ct,
            int workload) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            var raw = new List<CellRaw>(capacity: GetSnapshotCapacity(workload));
            SnapshotCellsInto(raw, r1, c1, r2, c2, ct);
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide(operationName, raw.Count);

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel && raw.Count > 0) {
                var po = new ParallelOptions {
                    CancellationToken = ct,
                    MaxDegreeOfParallelism = policy.MaxDegreeOfParallelism ?? -1
                };
                Parallel.For(0, raw.Count, po, i => raw[i] = ConvertRaw(raw[i]));
            } else {
                bool canCancel = ct.CanBeCanceled;
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                for (int i = 0; i < raw.Count; i++) {
                    if (canCancel && (i & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    raw[i] = ConvertRaw(raw[i]);
                }
            }

            return raw;
        }

        private static int GetSnapshotCapacity(int workload) {
            if (workload <= 0) {
                return 0;
            }

            if (workload <= DenseSnapshotCapacityLimit) {
                return workload;
            }

            return Math.Max(1024, workload / 4);
        }

        private void SnapshotCellsInto(List<CellRaw> buffer, int r1, int c1, int r2, int c2, CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    var raw = SnapshotCell(cell, rIndex, cIndex);

                    if (raw.RawText != null || raw.InlineText != null || raw.FormulaText != null || CellHasExplicitBlank(cell) || _opt.FillBlanksInRanges)
                        buffer.Add(raw);
                }
            }
        }

        private void FillRangeSequential(object?[,] result, int r1, int c1, int r2, int c2, CancellationToken ct) {
            int height = result.GetLength(0);
            int width = result.GetLength(1);
            bool canCancel = ct.CanBeCanceled;
            int visitedCells = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = checked((int)row.RowIndex!.Value);
                if (rIndex < r1) continue;
                if (rIndex > r2) continue;

                int rr = rIndex - r1;
                if ((uint)rr >= (uint)height) continue;

                foreach (var cell in row.Elements<Cell>()) {
                    if (canCancel && (++visitedCells & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int cIndex = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                    if (cIndex < c1 || cIndex > c2) continue;

                    int cc = cIndex - c1;
                    if ((uint)cc >= (uint)width) continue;

                    if (TryConvertCell(cell, out object? value))
                        result[rr, cc] = value;
                }
            }
        }
    }
}
