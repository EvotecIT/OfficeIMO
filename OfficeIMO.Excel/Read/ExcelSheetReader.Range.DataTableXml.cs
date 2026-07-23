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
        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "Inferred worksheet column types are normalized to OfficeIMO's closed scalar set and are used only as DataColumn conversion tokens.")]
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
            bool canUseOrderedFullWidthExit = canTrackColumns;
            int nextExpectedColumn = c1;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
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
                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

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

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                }

                if (canUseOrderedFullWidthExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }

                if (canTrackColumns && !canUseOrderedFullWidthExit && MarkRequestedColumnSeen(cc, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
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

        [UnconditionalSuppressMessage("Trimming", "IL2062", Justification = "Inferred worksheet column types are normalized to OfficeIMO's closed scalar set and are used only as DataColumn conversion tokens.")]
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
            long cellCount = (long)rows * cols;
            if (_opt.MaxRangeCells <= 0) {
                throw new ArgumentOutOfRangeException(nameof(_opt.MaxRangeCells), "Maximum dense range cell count must be positive.");
            }

            if (cellCount > _opt.MaxRangeCells) {
                throw new InvalidDataException(
                    $"Range '{a1Range}' contains {cellCount.ToString(CultureInfo.InvariantCulture)} cells, exceeding the configured limit of {_opt.MaxRangeCells.ToString(CultureInfo.InvariantCulture)}.");
            }

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

    }
}
