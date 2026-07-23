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
        /// <summary>
        /// Reads a rectangular range to a DataTable. If headersInFirstRow = true, first row becomes column names.
        /// </summary>
        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "Inferred worksheet column types are normalized to OfficeIMO's closed scalar set and are used only as DataColumn conversion tokens.")]
        public DataTable ReadRangeAsDataTable(string a1Range, bool headersInFirstRow = true, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) throw new ArgumentException($"Invalid range '{a1Range}'.");

            int rows = r2 - r1 + 1;
            int cols = c2 - c1 + 1;
            long cellCount = (long)rows * cols;
            if (_opt.MaxRangeCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxRangeCells), "Maximum dense range cell count must be positive.");
            }

            if (cellCount > _opt.MaxRangeCells) {
                throw new InvalidDataException(
                    $"Range '{a1Range}' contains {cellCount.ToString(CultureInfo.InvariantCulture)} cells, exceeding the configured limit of {_opt.MaxRangeCells.ToString(CultureInfo.InvariantCulture)}.");
            }

            var dt = new DataTable(_sheetName);
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            int workload = checked((int)cellCount);
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

        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "Inferred worksheet column types are normalized to OfficeIMO's closed scalar set and are used only as DataColumn conversion tokens.")]
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
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

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

            if (cols == 8) {
                return ReadXmlRowIntoDataTableBuffer8(rowReader, c1, c2, headerValues, rowValues, inferredTypes, ct);
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = cols <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(cols) : 0UL;
            ulong seenColumns = 0;
            bool canUseOrderedFullWidthExit = canTrackColumns;
            int nextExpectedColumn = c1;
            bool hasNullValue = false;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
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
                    if (canUseOrderedFullWidthExit) {
                        canUseOrderedFullWidthExit = false;
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
                        canUseOrderedFullWidthExit = false;
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int cc = columnIndex - c1;
                if ((uint)cc >= (uint)cols) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
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

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                }

                if (canUseOrderedFullWidthExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return !hasNullValue;
                }

                if (canTrackColumns && !canUseOrderedFullWidthExit && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    return !hasNullValue;
                }
            }

            return canTrackColumns && seenColumns == allColumnsSeen && !hasNullValue;
        }

        private bool ReadXmlRowIntoDataTableBuffer8(
            XmlReader rowReader,
            int c1,
            int c2,
            object?[]? headerValues,
            object?[]? rowValues,
            Type?[]? inferredTypes,
            CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
            bool hasNullValue = false;
            int visitedNodes = 0;

            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return seenColumns == 0xFFUL && !hasNullValue;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    if (canUseOrderedFullWidthExit) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    canUseOrderedFullWidthExit = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
                        int orderedSeen = nextExpectedColumn - c1;
                        seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                        canUseOrderedFullWidthExit = false;
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int columnOffset = columnIndex - c1;
                if ((uint)columnOffset >= 8U) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                }

                object? value = ReadXmlCellValue(rowReader);
                if (value == null) {
                    hasNullValue = true;
                }

                if (headerValues != null) {
                    headerValues[columnOffset] = value;
                } else if (rowValues != null) {
                    rowValues[columnOffset] = value;
                    if (inferredTypes != null && inferredTypes[columnOffset] != typeof(object)) {
                        inferredTypes[columnOffset] = MergeDataTableColumnType(inferredTypes[columnOffset], value);
                    }
                }

                seenColumns |= 1UL << columnOffset;
                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return !hasNullValue;
                    }
                } else if (seenColumns == 0xFFUL) {
                    SkipXmlElementContent(rowReader, depth);
                    return !hasNullValue;
                }
            }

            return seenColumns == 0xFFUL && !hasNullValue;
        }

    }
}
