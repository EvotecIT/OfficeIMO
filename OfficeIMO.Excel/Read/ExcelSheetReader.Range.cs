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
        private const int DataTableBufferedSinglePassCapacityLimit = 1_000_000;
        private const int XmlFastCompletedRowTrackingLimit = 4096;

        /// <summary>
        /// Returns the used range of the worksheet as an A1 string (e.g., "A1:C10").
        /// If the sheet is empty, returns "A1:A1".
        /// </summary>
        public string GetUsedRangeA1() {
            if (_usedRangeA1 != null) {
                return _usedRangeA1;
            }

            if (_canStreamWorksheetPart
                && TryGetWorksheetDimensionReferenceFromXml(out string dimensionReference)) {
                _usedRangeA1 = dimensionReference;
                return dimensionReference;
            }

            if (TryGetWorksheetDimensionReference(WorksheetRoot, out dimensionReference)) {
                _usedRangeA1 = dimensionReference;
                return dimensionReference;
            }

            string reference = ExcelSheet.ComputeSheetDimensionReference(WorksheetRoot);
            string usedRange = reference.IndexOf(":", StringComparison.Ordinal) >= 0 ? reference : reference + ":" + reference;
            _usedRangeA1 = usedRange;
            return usedRange;
        }

        private bool TryGetWorksheetDimensionReferenceFromXml(out string reference) {
            reference = string.Empty;
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) {
                        continue;
                    }

                    if (reader.LocalName == "dimension") {
                        return TryNormalizeWorksheetDimensionReference(reader.GetAttribute("ref"), out reference);
                    }

                    if (reader.LocalName == "sheetData") {
                        return false;
                    }
                }
            } catch (XmlException) {
                return false;
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            } catch (ObjectDisposedException) {
                return false;
            }

            return false;
        }

        private static bool TryGetWorksheetDimensionReference(Worksheet worksheet, out string reference) {
            reference = string.Empty;
            string? rawReference = worksheet.SheetDimension?.Reference?.Value;
            return TryNormalizeWorksheetDimensionReference(rawReference, out reference);
        }

        private static bool TryNormalizeWorksheetDimensionReference(string? rawReference, out string reference) {
            reference = string.Empty;
            if (string.IsNullOrWhiteSpace(rawReference)) {
                return false;
            }

            rawReference = rawReference!.Trim();
            if (rawReference.Equals("A1", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (rawReference.IndexOf(':') >= 0) {
                if (!A1.TryParseRange(rawReference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                    || firstRow <= 0
                    || firstColumn <= 0
                    || lastRow < firstRow
                    || lastColumn < firstColumn) {
                    return false;
                }

                reference = rawReference;
                return true;
            }

            if (!A1.TryParseCellReferenceFast(rawReference, out int row, out int column)
                || row <= 0
                || column <= 0) {
                return false;
            }

            reference = rawReference + ":" + rawReference;
            return true;
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
                if (CanUseAutomaticXmlReadFastPath(policy)) {
                    if (TryFillRangeXmlFast(result, r1, c1, r2, c2, ct)) {
                        return result;
                    }
                }

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
                if (CanUseAutomaticXmlReadFastPath(policy)) {
                    if (TryFillDataTableXmlBufferedSinglePass(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, workload, ct)) {
                        return dt;
                    }

                    if (TryFillDataTableXmlFast(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                        return dt;
                    }

                    if (TryFillDataTableSequentialSinglePass(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, ct)) {
                        return dt;
                    }
                }

                decided = policy.Decide("ReadRangeAsDataTable", workload);
            }

            if (decided == OfficeIMO.Excel.ExecutionMode.Sequential) {
                if (TryFillDataTableXmlBufferedSinglePass(dt, r1, c1, r2, c2, rows, cols, headersInFirstRow, workload, ct)) {
                    return dt;
                }

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

        private bool TryFillDataTableXmlBufferedSinglePass(
            DataTable dt,
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            bool headersInFirstRow,
            int workload,
            CancellationToken ct) {
            if (!_opt.InferDataTableColumnTypes
                || workload > DataTableBufferedSinglePassCapacityLimit
                || !CanUseDataTableXmlBufferedReader()) {
                return false;
            }

            int startRow = headersInFirstRow ? 1 : 0;
            int dataRowCount = Math.Max(0, rows - startRow);
            var headerValues = headersInFirstRow ? new object?[cols] : null;
            var rowValues = new object?[dataRowCount][];
            var completeRowsWithoutNulls = cols <= 64 ? new bool[dataRowCount] : null;
            var inferredTypes = new Type?[cols];

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(rows);
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
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (headersInFirstRow && rowIndex == r1) {
                        ReadXmlRowIntoDataTableBuffer(reader, c1, c2, cols, headerValues, null, null, ct);
                        seenRows.MarkSeen(0);
                        continue;
                    }

                    int rr = rowIndex - r1 - startRow;
                    if ((uint)rr >= (uint)dataRowCount) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    object?[] values = rowValues[rr] ??= new object?[cols];
                    if (ReadXmlRowIntoDataTableBuffer(reader, c1, c2, cols, null, values, inferredTypes, ct)) {
                        completeRowsWithoutNulls![rr] = true;
                    }

                    seenRows.MarkSeen(rowIndex - r1);
                }

                Type[] columnTypes = new Type[cols];
                for (int c = 0; c < cols; c++) {
                    columnTypes[c] = inferredTypes[c] ?? typeof(object);
                }

                if (headersInFirstRow && rows > 0) {
                    var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues?[c]?.ToString(), _opt.NormalizeHeaders);
                    for (int c = 0; c < cols; c++) {
                        dt.Columns.Add(headers[c], columnTypes[c]);
                    }
                } else {
                    for (int c = 0; c < cols; c++) {
                        dt.Columns.Add($"Column{c + 1}", columnTypes[c]);
                    }
                }

                dt.MinimumCapacity = Math.Max(dt.MinimumCapacity, dataRowCount);
                dt.BeginLoadData();
                try {
                    AddBufferedRowsToDataTable(dt, rowValues, completeRowsWithoutNulls, dataRowCount, cols, ct);
                } finally {
                    dt.EndLoadData();
                }

                return true;
            } catch (XmlException) {
                dt.Clear();
                dt.Columns.Clear();
                return false;
            } catch (IOException) {
                dt.Clear();
                dt.Columns.Clear();
                return false;
            } catch (UnauthorizedAccessException) {
                dt.Clear();
                dt.Columns.Clear();
                return false;
            } catch (ObjectDisposedException) {
                dt.Clear();
                dt.Columns.Clear();
                return false;
            }
        }

        private static void AddBufferedRowsToDataTable(DataTable dt, object?[][] rowValues, bool[]? completeRowsWithoutNulls, int dataRowCount, int cols, CancellationToken ct) {
            bool canCancel = ct.CanBeCanceled;
            object[]? blankRow = null;
            for (int r = 0; r < dataRowCount; r++) {
                if (canCancel && (r & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                object?[]? source = rowValues[r];
                if (source == null) {
                    blankRow ??= CreateDbNullRow(cols);
                    dt.Rows.Add(blankRow);
                    continue;
                }

                if (completeRowsWithoutNulls == null || !completeRowsWithoutNulls[r]) {
                    for (int c = 0; c < cols; c++) {
                        source[c] ??= DBNull.Value;
                    }
                }

                dt.Rows.Add(source);
            }
        }

        private static object[] CreateDbNullRow(int cols) {
            var values = new object[cols];
            for (int i = 0; i < values.Length; i++) {
                values[i] = DBNull.Value;
            }

            return values;
        }

        private bool ReadXmlRowIntoDataTableBuffer(
            XmlReader rowReader,
            int c1,
            int c2,
            int cols,
            object?[]? headerValues,
            object?[]? rowValues,
            Type?[]? inferredTypes,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return false;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = cols <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(cols) : 0UL;
            ulong seenColumns = 0;
            bool hasNullValue = false;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return canTrackColumns && seenColumns == allColumnsSeen && !hasNullValue;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                if ((uint)cc >= (uint)cols) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                object? value = ReadXmlCellValue(rowReader);
                if (value == null) {
                    hasNullValue = true;
                }

                if (headerValues != null) {
                    headerValues[cc] = value;
                } else if (rowValues != null) {
                    rowValues[cc] = value;
                    if (inferredTypes != null && inferredTypes[cc] != typeof(object)) {
                        inferredTypes[cc] = MergeDataTableColumnType(inferredTypes[cc], value);
                    }
                }

                if (canTrackColumns && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return !hasNullValue;
                }
            }

            return canTrackColumns && seenColumns == allColumnsSeen && !hasNullValue;
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
            if (headerValues == null && inferredTypes == null) {
                return true;
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int rowCount = r2 - r1 + 1;
                var seenRows = CreateCompletedRowTracker(rowCount);
                bool headerRead = !headersInFirstRow;
                int unresolvedInferredTypes = inferredTypes?.Length ?? 0;
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
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    bool isHeaderRow = headersInFirstRow && rowIndex == r1;
                    bool inferFromRow = inferredTypes != null && (!headersInFirstRow || rowIndex > r1);
                    if (!isHeaderRow && !inferFromRow) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlRowIntoDataTableMetadata(reader, c1, c2, headerValues, inferredTypes, isHeaderRow, inferFromRow, ct, ref unresolvedInferredTypes);
                    seenRows.MarkSeen(rowIndex - r1);
                    if (isHeaderRow) {
                        headerRead = true;
                    }

                    if (headerRead && (inferredTypes == null || unresolvedInferredTypes == 0)) {
                        return true;
                    }
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
            CancellationToken ct,
            ref int unresolvedInferredTypes) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int columnCount = c2 - c1 + 1;
            bool canTrackColumns = columnCount <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(columnCount) : 0UL;
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

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

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
                    Type? previous = inferredTypes[cc];
                    Type? inferred = MergeDataTableColumnType(previous, value);
                    inferredTypes[cc] = inferred;
                    if (previous != typeof(object) && inferred == typeof(object)) {
                        unresolvedInferredTypes--;
                    }
                }

                if (canTrackColumns && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
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

            var rowValues = new object?[dataRowCount][];
            var completeRowsWithoutNulls = cols <= 64 ? new bool[dataRowCount] : null;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(rows);
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
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (headersInFirstRow && rowIndex == r1) {
                        seenRows.MarkSeen(0);
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    int rr = rowIndex - r1 - startRow;
                    if ((uint)rr >= (uint)rowValues.Length) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    object?[] values = rowValues[rr] ??= new object?[cols];
                    if (ReadXmlRowIntoDataTableBuffer(reader, c1, c2, cols, null, values, null, ct)) {
                        completeRowsWithoutNulls![rr] = true;
                    }
                    seenRows.MarkSeen(rowIndex - r1);
                }

                dt.MinimumCapacity = Math.Max(dt.MinimumCapacity, dataRowCount);
                dt.BeginLoadData();
                try {
                    AddBufferedRowsToDataTable(dt, rowValues, completeRowsWithoutNulls, dataRowCount, cols, ct);
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

            if (CanUseReadObjectsXmlFastPath(mode)) {
                if (TryReadObjectsDictionaryXmlStreamingFast(r1, c1, r2, c2, rows, cols, ct, out var streamingResult)) {
                    return streamingResult;
                }

                if (TryReadObjectsDictionaryXmlFast(r1, c1, r2, c2, rows, cols, ct, out var xmlResult)) {
                    return xmlResult;
                }
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

        private bool CanUseReadObjectsXmlFastPath(OfficeIMO.Excel.ExecutionMode? mode) {
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel) {
                return false;
            }

            return policy.OnDecision == null && CanUseXmlFastReader();
        }

        private bool TryReadObjectsDictionaryXmlStreamingFast(
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
            var headerValues = new object?[cols];
            string[]? headers = null;
            int nextDataRow = r1 + 1;
            int lastDataRow = r1;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
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
                    if (rowIndex < r1) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex > r2) {
                        break;
                    }

                    if (rowIndex == r1) {
                        if (nextDataRow != r1 + 1 || result.Count > 0 || lastDataRow != r1) {
                            result = [];
                            return false;
                        }

                        ReadXmlRowValuesInto(reader, rowIndex, c1, c2, headerValues, ct);
                        headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                        continue;
                    }

                    if (rowIndex <= lastDataRow) {
                        result = [];
                        return false;
                    }

                    headers ??= ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                    while (nextDataRow < rowIndex) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        result.Add(CreateEmptyDictionaryRow(headers, cols));
                        nextDataRow++;
                    }

                    result.Add(ReadXmlRowIntoDictionary(reader, c1, c2, headers, cols, ct));
                    lastDataRow = rowIndex;
                    nextDataRow = rowIndex + 1;
                }

                headers ??= ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                while (nextDataRow <= r2) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    result.Add(CreateEmptyDictionaryRow(headers, cols));
                    nextDataRow++;
                }

                return result.Count == dataRowCount;
            } catch (XmlException) {
                result = [];
                return false;
            } catch (IOException) {
                result = [];
                return false;
            } catch (UnauthorizedAccessException) {
                result = [];
                return false;
            } catch (ObjectDisposedException) {
                result = [];
                return false;
            }
        }

        private Dictionary<string, object?> ReadXmlRowIntoDictionary(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return CreateEmptyDictionaryRow(headers, cols);
            }

            object?[] values = new object?[cols];
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = cols <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(cols) : 0UL;
            ulong seenColumns = 0;
            while (rowReader.Read()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return CreateDictionaryRow(headers, values, cols);
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int offset = columnIndex - c1;
                if ((uint)offset >= (uint)cols) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                values[offset] = ReadXmlCellValue(rowReader);
                if (canTrackColumns && MarkRequestedColumnSeen(offset, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return CreateDictionaryRow(headers, values, cols);
                }
            }

            return CreateDictionaryRow(headers, values, cols);
        }

        private static Dictionary<string, object?> CreateEmptyDictionaryRow(string[] headers, int columnCount) {
            var dict = new Dictionary<string, object?>(columnCount, StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < columnCount; c++) {
                dict.Add(headers[c], null);
            }

            return dict;
        }

        private static Dictionary<string, object?> CreateDictionaryRow(string[] headers, object?[]? values, int columnCount) {
            var dict = new Dictionary<string, object?>(columnCount, StringComparer.OrdinalIgnoreCase);
            if (values == null) {
                for (int c = 0; c < columnCount; c++) {
                    dict.Add(headers[c], null);
                }

                return dict;
            }

            for (int c = 0; c < columnCount; c++) {
                dict.Add(headers[c], values[c]);
            }

            return dict;
        }

        private bool TryReadObjectsDictionaryXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int rows,
            int cols,
            CancellationToken ct,
            out List<Dictionary<string, object?>> result) {
            result = [];
            int dataRowCount = rows - 1;
            var headerValues = new object?[cols];
            var rowValues = dataRowCount == 0 ? Array.Empty<object?[]>() : new object?[dataRowCount][];

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(rows);

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
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex == r1) {
                        ReadXmlRowValuesInto(reader, rowIndex, c1, c2, headerValues, ct);
                        seenRows.MarkSeen(0);
                        continue;
                    }

                    int rowOffset = rowIndex - r1 - 1;
                    if ((uint)rowOffset >= (uint)rowValues.Length) {
                        continue;
                    }

                    object?[] values = ReadXmlRowValues(reader, rowIndex, c1, c2, cols, ct);
                    object?[]? existing = rowValues[rowOffset];
                    if (existing == null) {
                        rowValues[rowOffset] = values;
                    } else {
                        MergeRowValues(existing, values);
                    }

                    seenRows.MarkSeen(rowIndex - r1);
                }

                var headers = ExcelHeaderNameHelper.BuildUniqueHeaders(cols, c => headerValues[c]?.ToString(), _opt.NormalizeHeaders);
                result = new List<Dictionary<string, object?>>(dataRowCount);
                for (int r = 0; r < dataRowCount; r++) {
                    if (canCancel && (r & 1023) == 0) {
                        ct.ThrowIfCancellationRequested();
                    }

                    result.Add(CreateDictionaryRow(headers, rowValues[r], cols));
                }

                return true;
            } catch (XmlException) {
                result = [];
                return false;
            } catch (IOException) {
                result = [];
                return false;
            } catch (UnauthorizedAccessException) {
                result = [];
                return false;
            } catch (ObjectDisposedException) {
                result = [];
                return false;
            }

            static void MergeRowValues(object?[] target, object?[] source) {
                for (int i = 0; i < source.Length; i++) {
                    if (source[i] != null) {
                        target[i] = source[i];
                    }
                }
            }

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
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int width = result.GetLength(1);
                int height = result.GetLength(0);
                var seenRows = CreateCompletedRowTracker(height);
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
                        if (rowIndex > r2 && seenRows.AllRowsSeen) {
                            break;
                        }

                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, ct);
                    seenRows.MarkSeen(rowIndex - r1);
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

        private bool CanUseAutomaticXmlReadFastPath(ExecutionPolicy policy) {
            return policy.OnDecision == null;
        }

        private bool CanUseDataTableXmlBufferedReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }

        private static CompletedRowTracker CreateCompletedRowTracker(int rowCount) {
            return new CompletedRowTracker(rowCount);
        }

        private struct CompletedRowTracker {
            private readonly int _rowCount;
            private bool[]? _seenRows;
            private ulong _seenRowMask0;
            private ulong _seenRowMask1;
            private ulong _seenRowMask2;
            private ulong _seenRowMask3;
            private int _seenRowCount;

            internal CompletedRowTracker(int rowCount) {
                if (rowCount <= 0 || rowCount > XmlFastCompletedRowTrackingLimit) {
                    _rowCount = 0;
                    _seenRows = null;
                    _seenRowMask0 = 0;
                    _seenRowMask1 = 0;
                    _seenRowMask2 = 0;
                    _seenRowMask3 = 0;
                    _seenRowCount = 0;
                    return;
                }

                _rowCount = rowCount;
                _seenRows = rowCount > 256 ? new bool[rowCount] : null;
                _seenRowMask0 = 0;
                _seenRowMask1 = 0;
                _seenRowMask2 = 0;
                _seenRowMask3 = 0;
                _seenRowCount = 0;
            }

            internal readonly bool AllRowsSeen => _rowCount > 0 && _seenRowCount == _rowCount;

            internal void MarkSeen(int rowOffset) {
                if ((uint)rowOffset >= (uint)_rowCount) {
                    return;
                }

                if (_seenRows == null) {
                    int maskIndex = rowOffset >> 6;
                    ulong rowBit = 1UL << (rowOffset & 63);
                    switch (maskIndex) {
                        case 0:
                            if ((_seenRowMask0 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask0 |= rowBit;
                            break;
                        case 1:
                            if ((_seenRowMask1 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask1 |= rowBit;
                            break;
                        case 2:
                            if ((_seenRowMask2 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask2 |= rowBit;
                            break;
                        default:
                            if ((_seenRowMask3 & rowBit) != 0) {
                                return;
                            }

                            _seenRowMask3 |= rowBit;
                            break;
                    }
                } else {
                    if (_seenRows[rowOffset]) {
                        return;
                    }

                    _seenRows[rowOffset] = true;
                }

                _seenRowCount++;
            }
        }

        private static ulong CreateAllColumnsSeenMask(int columnCount) {
            return columnCount == 64 ? ulong.MaxValue : (1UL << columnCount) - 1UL;
        }

        private static bool MarkRequestedColumnSeen(int columnOffset, ulong allColumnsSeen, ref ulong seenColumns) {
            seenColumns |= 1UL << columnOffset;
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
            bool canTrackColumns = width <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(width) : 0UL;
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

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

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
                if (canTrackColumns && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private object? ReadXmlCellValue(XmlReader cellReader) {
            if (cellReader.IsEmptyElement) {
                return null;
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            if (_opt.CellValueConverter != null) {
                CellRaw raw = ReadXmlCellRaw(cellReader, 0, 0, cellKind, readStyleIndex: true);
                return ConvertRaw(raw).TypedValue;
            }

            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            bool useDateStyle = _opt.TreatDatesUsingNumberFormat
                && _styles.HasDateStyles
                && CellKindCanUseDateStyle(cellKind);
            string? styleAttribute = useDateStyle ? cellReader.GetAttribute("s") : null;

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
                        if (useCachedFormulaResult) {
                            if (TryConvertXmlRawText(cellKind, rawText, useDateStyle, styleAttribute, numericAsDecimal, culture, out object? fastValue)) {
                                SkipXmlElementContent(cellReader, depth, "c");
                                return fastValue;
                            }
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
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

            if (formulaText != null && !useCachedFormulaResult) {
                return formulaText;
            }

            if (formulaText != null && rawText == null) {
                return formulaText;
            }

            if (cellKind == XmlCellKind.InlineString) {
                return inlineText;
            }

            if (cellKind == XmlCellKind.SharedString) {
                return TryParseSharedStringIndex(rawText, out int sstIndex) ? _sst.Get(sstIndex) : rawText;
            }

            if (cellKind == XmlCellKind.Boolean && rawText != null) {
                return rawText == "1";
            }

            if (cellKind == XmlCellKind.Date && rawText != null) {
                return DateTime.TryParse(rawText, culture, DateTimeStyles.AssumeLocal, out var date)
                    ? date
                    : rawText;
            }

            if (cellKind == XmlCellKind.String) {
                return rawText ?? inlineText;
            }

            if (rawText == null) {
                return inlineText;
            }

            if (useDateStyle
                && TryParseUInt(styleAttribute, out uint styleIndex)
                && _styles.IsDateLike(styleIndex)
                && (TryParseInvariantDoubleFast(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                return DateTime.FromOADate(oa);
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal decimalNumber)) {
                return decimalNumber;
            }

            return (TryParseInvariantDoubleFast(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
        }

        private bool TryConvertXmlRawText(
            XmlCellKind cellKind,
            string? rawText,
            bool useDateStyle,
            string? styleAttribute,
            bool numericAsDecimal,
            CultureInfo culture,
            out object? value) {
            value = null;
            if (rawText == null) {
                return false;
            }

            if (cellKind == XmlCellKind.SharedString) {
                value = TryParseSharedStringIndex(rawText, out int sstIndex) ? _sst.Get(sstIndex) : rawText;
                return true;
            }

            if (cellKind == XmlCellKind.Boolean) {
                value = rawText == "1";
                return true;
            }

            if (cellKind == XmlCellKind.Date) {
                value = DateTime.TryParse(rawText, culture, DateTimeStyles.AssumeLocal, out var date)
                    ? date
                    : rawText;
                return true;
            }

            if (cellKind == XmlCellKind.String) {
                value = rawText;
                return true;
            }

            if (cellKind == XmlCellKind.InlineString) {
                return false;
            }

            if (useDateStyle
                && TryParseUInt(styleAttribute, out uint styleIndex)
                && _styles.IsDateLike(styleIndex)
                && (TryParseInvariantDoubleFast(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                value = DateTime.FromOADate(oa);
                return true;
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal decimalNumber)) {
                value = decimalNumber;
                return true;
            }

            value = (TryParseInvariantDoubleFast(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
            return true;
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
