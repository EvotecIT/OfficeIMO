using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
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
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                decided = policy.Decide("ReadRange", workload);
            }

            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                if (TryFillRangeXmlFast(result, r1, c1, r2, c2, ct)) {
                    return result;
                }

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
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = rows * cols;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) {
                decided = policy.Decide("ReadRangeAsDataTable", workload);
            }

            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                if (TryFillDataTableXmlFast(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                    return dt;
                }

                if (TryFillDataTableSequentialSinglePass(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                    return dt;
                }

                FillDataTableSequential(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct);
                return dt;
            }

            var raw = SnapshotAndConvertRangeCells(r1, c1, r2, c2, "ReadRangeAsDataTable", decided, ct, workload);

            Type[]? columnTypes = _opt.InferDataTableColumnTypes
                ? InferDataTableColumnTypesFromRaw(raw, r1, c1, cols, headersInFirstRow ? 1 : 0)
                : null;

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
                    dt.Columns.Add(headers[c], columnTypes?[c] ?? typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) dt.Columns.Add($"Column{c + 1}", columnTypes?[c] ?? typeof(object));
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

        private bool TryFillDataTableXmlFast(
            DataTable dt,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            bool headersInFirstRow,
            CancellationToken ct) {
            if (!CanUseXmlFastReader()) {
                return false;
            }

            if (!TryReadDataTableXmlMetadata(r1, c1, r2, c2, cols, headersInFirstRow, ct, out var headerValues, out var columnTypes)) {
                return false;
            }

            if (headersInFirstRow && rows > 0) {
                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues?[c]?.ToString(), _opt.NormalizeHeaders);
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], columnTypes?[c] ?? typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add($"Column{c + 1}", columnTypes?[c] ?? typeof(object));
                }
            }

            if (TryFillDataTableRowsXmlFast(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                return true;
            }

            dt.Clear();
            dt.Columns.Clear();
            return false;
        }

        private bool TryReadDataTableXmlMetadata(
            int r1,
            int c1,
            int r2,
            int c2,
            int cols,
            bool headersInFirstRow,
            CancellationToken ct,
            out object?[]? headerValues,
            out Type[]? columnTypes) {
            headerValues = headersInFirstRow ? new object?[cols] : null;
            columnTypes = null;
            Type?[]? inferredTypes = _opt.InferDataTableColumnTypes ? new Type?[cols] : null;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1 || rowIndex > r2) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    bool isHeaderRow = headersInFirstRow && rowIndex == r1;
                    bool inferFromRow = inferredTypes != null && (!headersInFirstRow || rowIndex > r1);
                    if (!isHeaderRow && !inferFromRow) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlRowIntoDataTableMetadata(reader, c1, c2, headerValues, inferredTypes, isHeaderRow, inferFromRow, ct);
                }

                if (inferredTypes != null) {
                    columnTypes = new Type[cols];
                    for (int c = 0; c < cols; c++) {
                        columnTypes[c] = inferredTypes[c] ?? typeof(object);
                    }
                }

                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

        private void ReadXmlRowIntoDataTableMetadata(
            XmlReader rowReader,
            int c1,
            int c2,
            object?[]? headerValues,
            Type?[]? inferredTypes,
            bool isHeaderRow,
            bool inferFromRow,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int columnCount = c2 - c1 + 1;
            ulong seenColumns = 0;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                string? reference = rowReader.GetAttribute("r");
                int columnIndex = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(reference);
                if (columnIndex <= 0) {
                    if (!string.IsNullOrEmpty(reference)) {
                        SkipXmlElement(rowReader, "c");
                        continue;
                    }

                    columnIndex = nextColumnIndex;
                }

                nextColumnIndex = columnIndex + 1;
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                object? value = ReadXmlCellValue(rowReader);
                if (isHeaderRow) {
                    headerValues![cc] = value;
                }

                if (inferFromRow && inferredTypes![cc] != typeof(object)) {
                    inferredTypes[cc] = MergeDataTableColumnType(inferredTypes[cc], value);
                }

                if (MarkRequestedColumnSeen(cc, columnCount, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private bool TryFillDataTableRowsXmlFast(
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
            if (dataRowCount == 0) {
                return true;
            }

            var dataRows = new DataRow[dataRowCount];
            for (int r = 0; r < dataRowCount; r++) {
                dataRows[r] = dt.NewRow();
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1 || rowIndex > r2) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (headersInFirstRow && rowIndex == r1) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    int rr = rowIndex - r1 - startRow;
                    if ((uint)rr >= (uint)dataRows.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlRowIntoDataRow(reader, dataRows[rr], c1, c2, cols, ct);
                }

                dt.MinimumCapacity = Math.Max(dt.MinimumCapacity, dataRowCount);
                dt.BeginLoadData();
                try {
                    for (int r = 0; r < dataRows.Length; r++) {
                        if (canCancel && (r & 1023) == 0) {
                            ct.ThrowIfCancellationRequested();
                        }

                        dt.Rows.Add(dataRows[r]);
                    }
                } finally {
                    dt.EndLoadData();
                }

                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

        private void ReadXmlRowIntoDataRow(XmlReader rowReader, DataRow dataRow, int c1, int c2, int cols, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            ulong seenColumns = 0;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                string? reference = rowReader.GetAttribute("r");
                int columnIndex = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(reference);
                if (columnIndex <= 0) {
                    if (!string.IsNullOrEmpty(reference)) {
                        SkipXmlElement(rowReader, "c");
                        continue;
                    }

                    columnIndex = nextColumnIndex;
                }

                nextColumnIndex = columnIndex + 1;
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                if ((uint)cc >= (uint)cols) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                dataRow[cc] = ReadXmlCellValue(rowReader) ?? DBNull.Value;
                if (MarkRequestedColumnSeen(cc, cols, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private void FillDataTableFromMatrix(DataTable dt, object?[,] values, int rows, int cols, bool headersInFirstRow, CancellationToken ct) {
            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            Type[]? columnTypes = _opt.InferDataTableColumnTypes
                ? InferDataTableColumnTypes(values, startRow, rows, cols)
                : null;

            if (headersInFirstRow && rows > 0) {
                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => values[0, c]?.ToString(), _opt.NormalizeHeaders);
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add(headers[c], columnTypes?[c] ?? typeof(object));
                }
            } else {
                for (int c = 0; c < cols; c++) {
                    dt.Columns.Add($"Column{c + 1}", columnTypes?[c] ?? typeof(object));
                }
            }

            bool canCancel = ct.CanBeCanceled;
            dt.BeginLoadData();
            try {
                for (int r = 0; r < dataRowCount; r++) {
                    if (canCancel && (r & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int sourceRow = r + startRow;
                    var rowValues = new object?[cols];
                    for (int c = 0; c < cols; c++) {
                        rowValues[c] = values[sourceRow, c] ?? DBNull.Value;
                    }

                    dt.Rows.Add(rowValues);
                }
            } finally {
                dt.EndLoadData();
            }
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
            int inferredRowIndex = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rowIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
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

        private static Type[] InferDataTableColumnTypes(object?[][] rowValues, int cols) {
            var types = new Type[cols];
            for (int c = 0; c < cols; c++) {
                Type? inferred = null;
                for (int r = 0; r < rowValues.Length; r++) {
                    object?[]? row = rowValues[r];
                    if (row == null) {
                        continue;
                    }

                    inferred = MergeDataTableColumnType(inferred, row[c]);
                    if (inferred == typeof(object)) {
                        break;
                    }
                }

                types[c] = inferred ?? typeof(object);
            }

            return types;
        }

        private static Type[] InferDataTableColumnTypes(object?[,] values, int startRow, int rows, int cols) {
            var types = new Type[cols];
            for (int c = 0; c < cols; c++) {
                Type? inferred = null;
                for (int r = startRow; r < rows; r++) {
                    inferred = MergeDataTableColumnType(inferred, values[r, c]);
                    if (inferred == typeof(object)) {
                        break;
                    }
                }

                types[c] = inferred ?? typeof(object);
            }

            return types;
        }

        private static Type[] InferDataTableColumnTypesFromRaw(List<CellRaw> raw, int r1, int c1, int cols, int startRow) {
            var types = new Type[cols];
            Type?[] inferred = new Type?[cols];
            for (int i = 0; i < raw.Count; i++) {
                var cell = raw[i];
                int rr = cell.Row - r1 - startRow;
                int cc = cell.Col - c1;
                if (rr < 0 || (uint)cc >= (uint)cols || inferred[cc] == typeof(object)) {
                    continue;
                }

                inferred[cc] = MergeDataTableColumnType(inferred[cc], cell.TypedValue);
            }

            for (int c = 0; c < cols; c++) {
                types[c] = inferred[c] ?? typeof(object);
            }

            return types;
        }

        private static Type? MergeDataTableColumnType(Type? current, object? value) {
            if (value == null || value == DBNull.Value) {
                return current;
            }

            Type next = value.GetType();
            if (current == null || current == next) {
                return next;
            }

            return typeof(object);
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
            int inferredRowIndex = 0;
            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
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
            int inferredRowIndex = 0;

            foreach (var row in EnumerateWorksheetRows(ct)) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                var rIndex = GetSequentialRowIndex(row, ref inferredRowIndex);
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

        private static int GetSequentialRowIndex(Row row, ref int inferredRowIndex) {
            if (row.RowIndex != null) {
                inferredRowIndex = checked((int)row.RowIndex.Value);
            } else {
                inferredRowIndex++;
            }

            return inferredRowIndex;
        }

        private bool TryFillRangeXmlFast(object?[,] result, int r1, int c1, int r2, int c2, CancellationToken ct) {
            if (!CanUseXmlFastReader()) {
                return false;
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                var settings = new XmlReaderSettings {
                    DtdProcessing = DtdProcessing.Prohibit,
                    IgnoreComments = true,
                    IgnoreProcessingInstructions = true,
                    IgnoreWhitespace = true,
                    CloseInput = false
                };

                using var reader = XmlReader.Create(stream, settings);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int width = result.GetLength(1);
                while (reader.Read()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < r1 || rowIndex > r2) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, ct);
                }

                return true;
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }
        }

        private bool CanUseXmlFastReader() {
            return _opt.CellValueConverter == null
                && _opt.Culture == CultureInfo.InvariantCulture
                && CanStreamWorksheetPart();
        }

        private static bool MarkRequestedColumnSeen(int columnOffset, int columnCount, ref ulong seenColumns) {
            if ((uint)columnOffset >= (uint)columnCount || (uint)columnCount > 64u) {
                return false;
            }

            seenColumns |= 1UL << columnOffset;
            ulong allColumnsSeen = columnCount == 64 ? ulong.MaxValue : (1UL << columnCount) - 1UL;
            return seenColumns == allColumnsSeen;
        }

        private void ReadXmlRowIntoRange(XmlReader rowReader, object?[,] result, int rowIndex, int r1, int c1, int c2, int width, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int rr = rowIndex - r1;
            if ((uint)rr >= (uint)result.GetLength(0)) {
                SkipXmlElement(rowReader, "row");
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            ulong seenColumns = 0;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                string? reference = rowReader.GetAttribute("r");
                int columnIndex = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(reference);
                if (columnIndex <= 0) {
                    if (!string.IsNullOrEmpty(reference)) {
                        SkipXmlElement(rowReader, "c");
                        continue;
                    }

                    columnIndex = nextColumnIndex;
                }

                nextColumnIndex = columnIndex + 1;
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                if ((uint)cc >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                result[rr, cc] = ReadXmlCellValue(rowReader);
                if (MarkRequestedColumnSeen(cc, width, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private object? ReadXmlCellValue(XmlReader cellReader) {
            string? type = cellReader.GetAttribute("t");
            string? styleAttribute = _opt.TreatDatesUsingNumberFormat
                && type != "s"
                && type != "b"
                && type != "inlineStr"
                && type != "d"
                && type != "str"
                ? cellReader.GetAttribute("s")
                : null;

            if (cellReader.IsEmptyElement) {
                return _opt.FillBlanksInRanges ? null : null;
            }

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        rawText = cellReader.ReadElementContentAsString();
                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!_opt.UseCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth, "c");
                            return formulaText;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        inlineText = ReadXmlInlineString(cellReader);
                        hasNode = true;
                        continue;
                    }
                }

                hasNode = cellReader.Read();
            }

            if (formulaText != null && !_opt.UseCachedFormulaResult) {
                return formulaText;
            }

            if (formulaText != null && rawText == null) {
                return formulaText;
            }

            if (type == "inlineStr") {
                return inlineText;
            }

            if (type == "s") {
                return TryParseSharedStringIndex(rawText, out int sstIndex) ? _sst.Get(sstIndex) : rawText;
            }

            if (type == "b" && rawText != null) {
                return rawText == "1";
            }

            if (type == "d" && rawText != null) {
                return DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var date)
                    ? date
                    : rawText;
            }

            if (type == "str") {
                return rawText ?? inlineText;
            }

            if (rawText == null) {
                return inlineText;
            }

            if (_opt.TreatDatesUsingNumberFormat
                && TryParseUInt(styleAttribute, out uint styleIndex)
                && _styles.IsDateLike(styleIndex)
                && (TryParseInvariantDoubleFast(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                return DateTime.FromOADate(oa);
            }

            if (_opt.NumericAsDecimal
                && decimal.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, _opt.Culture, out decimal decimalNumber)) {
                return decimalNumber;
            }

            return (TryParseInvariantDoubleFast(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
        }

        private static string ReadXmlInlineString(XmlReader inlineReader) {
            if (inlineReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = inlineReader.Depth;
            string? first = null;
            System.Text.StringBuilder? builder = null;
            while (inlineReader.Read()) {
                if (inlineReader.NodeType == XmlNodeType.EndElement && inlineReader.Depth == depth && inlineReader.LocalName == "is") {
                    break;
                }

                if (inlineReader.NodeType != XmlNodeType.Element || inlineReader.LocalName != "t") {
                    continue;
                }

                string text = inlineReader.ReadElementContentAsString();
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new System.Text.StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        private static int ParsePositiveIntAttribute(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return 0;
            }

            string text = value!;
            int result = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return 0;
                }

                if (result > (int.MaxValue - digit) / 10) {
                    return 0;
                }

                result = (result * 10) + digit;
            }

            return result;
        }

        private static bool TryParseUInt(string? value, out uint result) {
            result = 0;
            if (string.IsNullOrEmpty(value)) {
                return false;
            }

            string text = value!;
            uint parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                uint digit = (uint)(text[i] - '0');
                if (digit > 9U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                if (parsed > (uint.MaxValue - digit) / 10U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                parsed = (parsed * 10U) + digit;
            }

            result = parsed;
            return true;
        }
    }
}
