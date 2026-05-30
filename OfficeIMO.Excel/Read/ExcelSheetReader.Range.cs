using DocumentFormat.OpenXml.Packaging;
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
        // DataTable materialization already carries row storage cost; keep the single-pass XML buffer active
        // through larger normal sheets to avoid a slower second worksheet scan.
        private const int DataTableBufferedSinglePassCapacityLimit = 3_000_000;
        private const int SparseReadInitialBufferCapacity = 64;
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
                && TryGetWorksheetDimensionReferenceFromXml(out string dimensionReference)
                && TryGetTableBackedDimensionReference(dimensionReference, out string tableBackedReference)) {
                _usedRangeA1 = tableBackedReference;
                return tableBackedReference;
            }

            if (_canStreamWorksheetPart
                && TryComputeUsedRangeReferenceFromXml(out string usedRangeReference)) {
                _usedRangeA1 = usedRangeReference;
                return usedRangeReference;
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
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

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

        private bool TryComputeUsedRangeReferenceFromXml(out string reference) {
            reference = string.Empty;
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    return false;
                }

                using var reader = OpenWorksheetXmlReader(stream);

                int minRow = int.MaxValue;
                int minColumn = int.MaxValue;
                int maxRow = 0;
                int maxColumn = 0;
                int nextRowIndex = 1;

                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    bool hasExplicitRowIndex = rowIndex > 0;
                    if (!hasExplicitRowIndex) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (reader.IsEmptyElement) {
                        continue;
                    }

                    int rowMinRow = hasExplicitRowIndex ? rowIndex : int.MaxValue;
                    int rowMaxRow = hasExplicitRowIndex ? rowIndex : 0;
                    int rowMinColumn = int.MaxValue;
                    int rowMaxColumn = 0;
                    int rowDepth = reader.Depth;
                    int lastColumn = 0;
                    bool advanceReader = true;
                    while (advanceReader ? reader.Read() : !reader.EOF) {
                        advanceReader = true;
                        if (reader.NodeType == XmlNodeType.EndElement
                            && reader.Depth == rowDepth
                            && reader.LocalName == "row") {
                            break;
                        }

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c") {
                            continue;
                        }

                        int column = 0;
                        string? cellReference = reader.GetAttribute("r");
                        if (hasExplicitRowIndex) {
                            column = A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(cellReference);
                        } else if (A1.TryParseCellReferenceFast(cellReference, out int parsedRow, out int parsedColumn)) {
                            column = parsedColumn;
                            if (parsedRow > 0) {
                                if (parsedRow < rowMinRow) rowMinRow = parsedRow;
                                if (parsedRow > rowMaxRow) rowMaxRow = parsedRow;
                            }
                        }

                        if (column <= 0) {
                            column = lastColumn + 1;
                        }

                        lastColumn = column;

                        if (column > 0) {
                            if (column < rowMinColumn) rowMinColumn = column;
                            if (column > rowMaxColumn) rowMaxColumn = column;
                        }

                        if (!reader.IsEmptyElement) {
                            reader.Skip();
                            advanceReader = false;
                        }
                    }

                    if (rowMaxColumn <= 0) {
                        continue;
                    }

                    if (rowMaxRow <= 0) {
                        rowMinRow = rowIndex;
                        rowMaxRow = rowIndex;
                    }

                    if (rowMinRow < minRow) minRow = rowMinRow;
                    if (rowMaxRow > maxRow) maxRow = rowMaxRow;
                    if (rowMinColumn < minColumn) minColumn = rowMinColumn;
                    if (rowMaxColumn > maxColumn) maxColumn = rowMaxColumn;
                    if (!hasExplicitRowIndex) {
                        nextRowIndex = rowMaxRow + 1;
                    }
                }

                if (maxRow <= 0 || maxColumn <= 0) {
                    return false;
                }

                reference = A1.CellReference(minRow, minColumn) + ":" + A1.CellReference(maxRow, maxColumn);
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

        private static bool TryGetWorksheetDimensionReference(Worksheet worksheet, out string reference) {
            reference = string.Empty;
            string? rawReference = worksheet.SheetDimension?.Reference?.Value;
            return TryNormalizeWorksheetDimensionReference(rawReference, out reference);
        }

        private bool TryGetTableBackedDimensionReference(string dimensionReference, out string reference) {
            reference = string.Empty;
            if (!A1.TryParseRange(dimensionReference, out int dimensionFirstRow, out int dimensionFirstColumn, out int dimensionLastRow, out int dimensionLastColumn)) {
                return false;
            }

            int minRow = int.MaxValue;
            int minColumn = int.MaxValue;
            int maxRow = 0;
            int maxColumn = 0;

            try {
                foreach (var tablePart in _wsPart.TableDefinitionParts) {
                    string? tableReference = TryGetTableReferenceXmlFast(tablePart, out string xmlTableReference)
                        ? xmlTableReference
                        : tablePart.Table?.Reference?.Value;
                    if (string.IsNullOrWhiteSpace(tableReference)
                        || !A1.TryParseRange(tableReference!, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                        return false;
                    }

                    if (firstRow < minRow) minRow = firstRow;
                    if (firstColumn < minColumn) minColumn = firstColumn;
                    if (lastRow > maxRow) maxRow = lastRow;
                    if (lastColumn > maxColumn) maxColumn = lastColumn;
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

            if (maxRow <= 0
                || minRow != dimensionFirstRow
                || minColumn != dimensionFirstColumn
                || maxRow != dimensionLastRow
                || maxColumn != dimensionLastColumn) {
                return false;
            }

            reference = dimensionReference;
            return true;
        }

        private static bool TryGetTableReferenceXmlFast(TableDefinitionPart tablePart, out string reference) {
            reference = string.Empty;
            try {
                using var stream = tablePart.GetStream(FileMode.Open, FileAccess.Read);
                using var reader = XmlReader.Create(stream, WorksheetXmlReaderSettings);
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "table") {
                        continue;
                    }

                    string? tableReference = reader.GetAttribute("ref");
                    if (string.IsNullOrWhiteSpace(tableReference)) {
                        return false;
                    }

                    reference = tableReference!;
                    return true;
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

            if (cols > 64) {
                return ReadXmlRowIntoDictionaryBuffered(rowReader, c1, c2, headers, cols, ct);
            }

            Dictionary<string, object?>? orderedDictionary = null;
            if (TryReadXmlOrderedFullWidthRowIntoDictionary(rowReader, c1, c2, headers, cols, ct, out orderedDictionary)) {
                return orderedDictionary ?? CreateEmptyDictionaryRow(headers, cols);
            }

            return orderedDictionary ?? CreateEmptyDictionaryRow(headers, cols);
        }

        private bool TryReadXmlOrderedFullWidthRowIntoDictionary(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct,
            out Dictionary<string, object?>? result) {
            result = new Dictionary<string, object?>(cols, StringComparer.OrdinalIgnoreCase);
            object?[] values = new object?[cols];
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            ulong allColumnsSeen = CreateAllColumnsSeenMask(cols);
            ulong seenColumns = 0;
            int nextExpectedColumn = c1;
            bool orderedFullWidth = true;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    result = orderedFullWidth && seenColumns == allColumnsSeen
                        ? result
                        : CreateDictionaryRow(headers, values, cols);
                    return true;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex <= 0) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                int offset = columnIndex - c1;
                if ((uint)offset >= (uint)cols) {
                    orderedFullWidth = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (orderedFullWidth && columnIndex != nextExpectedColumn) {
                    orderedFullWidth = false;
                }

                object? value = ReadXmlCellValue(rowReader);
                values[offset] = value;
                if (orderedFullWidth) {
                    result.Add(headers[offset], value);
                    nextExpectedColumn++;
                }

                if (orderedFullWidth && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return true;
                }

                if (!orderedFullWidth && MarkRequestedColumnSeen(offset, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    result = CreateDictionaryRow(headers, values, cols);
                    return true;
                }
            }

            result = orderedFullWidth && seenColumns == allColumnsSeen
                ? result
                : CreateDictionaryRow(headers, values, cols);
            return true;
        }

        private Dictionary<string, object?> ReadXmlRowIntoDictionaryBuffered(
            XmlReader rowReader,
            int c1,
            int c2,
            string[] headers,
            int cols,
            CancellationToken ct) {
            object?[] values = new object?[cols];
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return CreateDictionaryRow(headers, values, cols);
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
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
            SnapshotCellsInto(raw, r1, c1, r2, c2, ct, out bool needsSharedStrings, out bool needsStyles);
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic) decided = policy.Decide(operationName, raw.Count);

            if (decided == OfficeIMO.Excel.ExecutionMode.Parallel && raw.Count > 0) {
                PrepareCachesForParallelConversion(needsSharedStrings, needsStyles);
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

        private void PrepareCachesForParallelConversion(bool needsSharedStrings, bool needsStyles) {
            if (needsSharedStrings) {
                _sst.EnsureLoaded();
            }

            if (needsStyles) {
                _ = Styles;
            }
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

        private void SnapshotCellsInto(List<CellRaw> buffer, int r1, int c1, int r2, int c2, CancellationToken ct, out bool needsSharedStrings, out bool needsStyles) {
            needsSharedStrings = false;
            needsStyles = false;
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

                    if (raw.RawText != null || raw.InlineText != null || raw.FormulaText != null || CellHasExplicitBlank(cell) || _opt.FillBlanksInRanges) {
                        buffer.Add(raw);
                        if (!needsSharedStrings && raw.TypeHint == CellValues.SharedString) {
                            needsSharedStrings = true;
                        }

                        if (!needsStyles && _opt.TreatDatesUsingNumberFormat && raw.StyleIndex is not null) {
                            needsStyles = true;
                        }
                    }
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
            if (!CanAttemptXmlFastReader()) {
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
                bool orderedRows = true;
                int orderedRowsSeen = 0;
                if (canCancel) {
                    while (reader.Read()) {
                        ct.ThrowIfCancellationRequested();

                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == height : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, ct);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == height) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                } else {
                    while (reader.Read()) {
                        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                            continue;
                        }

                        int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                        if (rowIndex <= 0) {
                            rowIndex = nextRowIndex;
                        }

                        nextRowIndex = rowIndex + 1;
                        if (rowIndex < r1 || rowIndex > r2) {
                            bool allRowsSeen = orderedRows ? orderedRowsSeen == height : seenRows.AllRowsSeen;
                            if (rowIndex > r2 && allRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        ReadXmlRowIntoRange(reader, result, rowIndex, r1, c1, c2, width, CancellationToken.None);
                        if (orderedRows && rowIndex == r1 + orderedRowsSeen) {
                            orderedRowsSeen++;
                            if (orderedRowsSeen == height) {
                                break;
                            }

                            continue;
                        }

                        if (orderedRows) {
                            for (int row = 0; row < orderedRowsSeen; row++) {
                                seenRows.MarkSeen(row);
                            }

                            orderedRows = false;
                        }

                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
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

        private bool CanUseXmlFastReader() {
            return _opt.CellValueConverter == null
                && _opt.Culture == CultureInfo.InvariantCulture
                && CanStreamWorksheetPart();
        }

        private bool CanAttemptXmlFastReader() {
            return _opt.CellValueConverter == null
                && _opt.Culture == CultureInfo.InvariantCulture
                && _canStreamWorksheetPart;
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
                _seenRows = null;
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

                if (_seenRows == null
                    && _seenRowMask0 == 0
                    && _seenRowMask1 == 0
                    && _seenRowMask2 == 0
                    && _seenRowMask3 == 0) {
                    if (rowOffset < _seenRowCount) {
                        return;
                    }

                    if (rowOffset == _seenRowCount) {
                        _seenRowCount++;
                        return;
                    }

                    if (_rowCount > 256) {
                        _seenRows = CreateSeenRowsTracker(_seenRowCount, _rowCount);
                    } else {
                        MarkDensePrefixSeenInMasks(_seenRowCount, ref _seenRowMask0, ref _seenRowMask1, ref _seenRowMask2, ref _seenRowMask3);
                    }
                }

                if (_rowCount > 256) {
                    if (_seenRows == null) {
                        if (rowOffset < _seenRowCount) {
                            return;
                        }

                        if (rowOffset == _seenRowCount) {
                            _seenRowCount++;
                            return;
                        }

                        _seenRows = CreateSeenRowsTracker(_seenRowCount, _rowCount);
                    }

                    if (_seenRows[rowOffset]) {
                        return;
                    }

                    _seenRows[rowOffset] = true;
                    _seenRowCount++;
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

        private static void MarkDensePrefixSeenInMasks(int seenDensePrefixLength, ref ulong mask0, ref ulong mask1, ref ulong mask2, ref ulong mask3) {
            if (seenDensePrefixLength <= 0) {
                return;
            }

            if (seenDensePrefixLength >= 64) {
                mask0 = ulong.MaxValue;
            } else {
                mask0 = (1UL << seenDensePrefixLength) - 1UL;
                return;
            }

            int remaining = seenDensePrefixLength - 64;
            if (remaining <= 0) {
                return;
            }

            if (remaining >= 64) {
                mask1 = ulong.MaxValue;
            } else {
                mask1 = (1UL << remaining) - 1UL;
                return;
            }

            remaining -= 64;
            if (remaining <= 0) {
                return;
            }

            if (remaining >= 64) {
                mask2 = ulong.MaxValue;
            } else {
                mask2 = (1UL << remaining) - 1UL;
                return;
            }

            remaining -= 64;
            if (remaining > 0) {
                mask3 = remaining >= 64 ? ulong.MaxValue : (1UL << remaining) - 1UL;
            }
        }

        private static bool[] CreateSeenRowsTracker(int seenDensePrefixLength, int rowCount) {
            var seenRows = new bool[rowCount];
            for (int i = 0; i < seenDensePrefixLength; i++) {
                seenRows[i] = true;
            }

            return seenRows;
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

            if (width == 8) {
                ReadXmlRowIntoRange8(rowReader, result, rr, c1, c2, ct);
                return;
            }

            if (width == 3) {
                ReadXmlRowIntoRange3(rowReader, result, rr, c1, c2, ct);
                return;
            }

            if (width == 10) {
                ReadXmlRowIntoRangeKnownWidth(rowReader, result, rr, c1, c2, width, 0x3FFUL, ct);
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = width <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(width) : 0UL;
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
                if ((uint)cc >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

                result[rr, cc] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));
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

        private void ReadXmlRowIntoRange8(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
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
                    canUseOrderedFullWidthExit = false;
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedFullWidthExit && columnIndex > c2 && nextExpectedColumn <= c2) {
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
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                result[rowOffset, columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == 0xFFUL) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

        private void ReadXmlRowIntoRange3(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
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
                if ((uint)columnOffset >= 3U) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                result[rowOffset, columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == 0x7UL) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

        private void ReadXmlRowIntoRangeKnownWidth(XmlReader rowReader, object?[,] result, int rowOffset, int c1, int c2, int width, ulong allColumnsSeen, CancellationToken ct) {
            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int nextExpectedColumn = c1;
            bool canUseOrderedFullWidthExit = true;
            ulong seenColumns = 0;
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
                if ((uint)columnOffset >= (uint)width) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    canUseOrderedFullWidthExit = false;
                }

                result[rowOffset, columnOffset] = ReadXmlCellValue(rowReader, rowReader.GetAttribute("t"));

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                } else {
                    seenColumns |= 1UL << columnOffset;
                    if (seenColumns == allColumnsSeen) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

        private object? ReadXmlCellValue(XmlReader cellReader) {
            return ReadXmlCellValue(cellReader, cellReader.GetAttribute("t"));
        }

        private object? ReadXmlCellValue(XmlReader cellReader, string? cellType) {
            if (cellReader.IsEmptyElement) {
                return null;
            }

            if (_opt.CellValueConverter == null && cellType == "s") {
                var sharedStringItems = _sharedStringItems ??= _sst.GetItems();
                return ReadXmlSharedStringCellValue(cellReader, _opt.UseCachedFormulaResult, sharedStringItems);
            }

            if (_opt.CellValueConverter == null
                && (string.IsNullOrEmpty(cellType) || cellType == "n")) {
                return ReadXmlNumericCellValue(cellReader);
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellType);
            if (_opt.CellValueConverter != null) {
                CellRaw raw = ReadXmlCellRaw(cellReader, 0, 0, cellKind, readStyleIndex: true);
                return ConvertRaw(raw).TypedValue;
            }

            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            if (cellKind == XmlCellKind.SharedString) {
                var sharedStringItems = _sharedStringItems ??= _sst.GetItems();
                return ReadXmlSharedStringCellValue(cellReader, useCachedFormulaResult, sharedStringItems);
            }

            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            bool useDateStyle = false;
            if (_opt.TreatDatesUsingNumberFormat && CellKindCanUseDateStyle(cellKind)) {
                string? styleAttribute = cellReader.GetAttribute("s");
                useDateStyle = IsDateStyleAttribute(styleAttribute);
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
                        if (useCachedFormulaResult) {
                            rawText = ReadXmlValueTextAndSkipCell(cellReader, depth);
                        } else {
                            rawText = ReadXmlValueText(cellReader);
                        }

                        if (useCachedFormulaResult) {
                            if (!numericAsDecimal
                                && !useDateStyle
                                && (cellKind == XmlCellKind.Default || cellKind == XmlCellKind.Number)
                                && (TryParseInvariantDoubleFast(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (TryConvertXmlRawText(cellKind, rawText, useDateStyle, numericAsDecimal, culture, out object? fastValue)) {
                                return fastValue;
                            }
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
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
                return TryParseSharedStringIndex(rawText, out int sstIndex) ? GetSharedString(sstIndex) : rawText;
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

        private object? ReadXmlNumericCellValue(XmlReader cellReader) {
            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            bool useDateStyle = _opt.TreatDatesUsingNumberFormat && IsDateStyleAttribute(cellReader.GetAttribute("s"));

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
                        rawText = useCachedFormulaResult
                            ? ReadXmlValueTextAndSkipCell(cellReader, depth)
                            : ReadXmlValueText(cellReader);

                        if (useCachedFormulaResult) {
                            if (!numericAsDecimal
                                && !useDateStyle
                                && (TryParseInvariantDoubleFast(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (useDateStyle
                                && (TryParseInvariantDoubleFast(rawText, out double oa)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                                return DateTime.FromOADate(oa);
                            }

                            if (rawText == null) {
                                return null;
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

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
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

            if (rawText == null) {
                return inlineText;
            }

            if (useDateStyle
                && (TryParseInvariantDoubleFast(rawText, out double oaValue)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oaValue))) {
                return DateTime.FromOADate(oaValue);
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal rawDecimalNumber)) {
                return rawDecimalNumber;
            }

            return (TryParseInvariantDoubleFast(rawText, out double rawNumber)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out rawNumber))
                ? rawNumber
                : rawText;
        }

        private bool TryConvertXmlRawText(
            XmlCellKind cellKind,
            string? rawText,
            bool useDateStyle,
            bool numericAsDecimal,
            CultureInfo culture,
            out object? value) {
            value = null;
            if (rawText == null) {
                return false;
            }

            switch (cellKind) {
                case XmlCellKind.SharedString:
                    value = TryParseSharedStringIndex(rawText, out int sstIndex) ? GetSharedString(sstIndex) : rawText;
                    return true;
                case XmlCellKind.Boolean:
                    value = rawText == "1";
                    return true;
                case XmlCellKind.Date:
                    value = DateTime.TryParse(rawText, culture, DateTimeStyles.AssumeLocal, out var date)
                        ? date
                        : rawText;
                    return true;
                case XmlCellKind.String:
                    value = rawText;
                    return true;
                case XmlCellKind.InlineString:
                    return false;
            }

            if (useDateStyle
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

        private object? ReadXmlSharedStringCellValue(XmlReader cellReader, bool useCachedFormulaResult, List<string> sharedStringItems) {
            int depth = cellReader.Depth;
            string? rawText = null;
            string? formulaText = null;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        bool parsedSharedStringIndex = useCachedFormulaResult
                            ? TryReadXmlSharedStringIndexValueAndSkipCell(cellReader, depth, out int sstIndex, out rawText)
                            : TryReadXmlSharedStringIndexValue(cellReader, out sstIndex, out rawText);
                        if (useCachedFormulaResult) {
                            return parsedSharedStringIndex ? GetSharedString(sstIndex, sharedStringItems) : rawText;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            return formulaText;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        _ = ReadXmlInlineString(cellReader);
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

            return TryParseSharedStringIndex(rawText, out int index) ? GetSharedString(index, sharedStringItems) : rawText;
        }

        private static string? ReadXmlValueText(XmlReader valueReader) {
            if (valueReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = valueReader.Depth;
            if (!valueReader.Read()) {
                return null;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, depth);
                return null;
            }

            string text = valueReader.Value;
            SkipXmlElementContent(valueReader, depth);
            return text;
        }

        private static string? ReadXmlValueTextAndSkipCell(XmlReader valueReader, int cellDepth) {
            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                return string.Empty;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return null;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, cellDepth);
                return null;
            }

            string text = valueReader.Value;
            if (valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth) {
                return text;
            }

            SkipXmlElementContent(valueReader, cellDepth);
            return text;
        }

        private static bool TryReadXmlSharedStringIndexValue(XmlReader valueReader, out int index, out string? rawText) {
            index = 0;
            rawText = null;

            if (valueReader.IsEmptyElement) {
                rawText = string.Empty;
                return false;
            }

            int depth = valueReader.Depth;
            if (!valueReader.Read()) {
                return false;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, depth);
                return false;
            }

            string text = valueReader.Value;
            rawText = text;
            int parsed = 0;
            bool hasDigit = false;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    SkipXmlElementContent(valueReader, depth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    SkipXmlElementContent(valueReader, depth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
                hasDigit = true;
            }

            SkipXmlElementContent(valueReader, depth);
            if (!hasDigit) {
                return false;
            }

            index = parsed;
            return true;
        }

        private static bool TryReadXmlSharedStringIndexValueAndSkipCell(XmlReader valueReader, int cellDepth, out int index, out string? rawText) {
            index = 0;
            rawText = null;

            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                rawText = string.Empty;
                return false;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return false;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, cellDepth);
                return false;
            }

            string text = valueReader.Value;
            rawText = text;
            int parsed = 0;
            bool hasDigit = false;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    SkipXmlElementContent(valueReader, cellDepth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    SkipXmlElementContent(valueReader, cellDepth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
                hasDigit = true;
            }

            if (valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth) {
                if (!hasDigit) {
                    return false;
                }

                index = parsed;
                return true;
            }

            SkipXmlElementContent(valueReader, cellDepth);
            if (!hasDigit) {
                return false;
            }

            index = parsed;
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
