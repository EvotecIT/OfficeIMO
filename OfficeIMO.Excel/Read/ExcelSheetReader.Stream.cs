using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private const int BufferedRangeStreamRowLimit = 4_096;
        private const int OrderedBufferedRangeStreamCellLimit = 1_000_000;

        /// <summary>
        /// Lazily reads a rectangular A1 range as ordered row chunks. DOM traversal is single-threaded;
        /// per-chunk value conversion is offloaded in parallel based on Execution policy.
        /// </summary>
        public IEnumerable<RangeChunk> ReadRangeStream(string a1Range, int chunkRows = 1024, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default) {
            if (chunkRows <= 0) throw new ArgumentOutOfRangeException(nameof(chunkRows), "Chunk row count must be greater than zero.");

            (int r1, int c1, int r2, int c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) yield break;

            bool canCancel = ct.CanBeCanceled;
            if (canCancel) {
                ct.ThrowIfCancellationRequested();
            }

            int estRows = Math.Max(0, r2 - r1 + 1);
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            bool automaticDecision = decided == OfficeIMO.Excel.ExecutionMode.Automatic;
            if (automaticDecision) {
                decided = policy.Decide("ReadRangeStream", estRows);
            }

            if (CanUseAutomaticXmlStreamFastPath(automaticDecision, decided)) {
                if (chunkRows >= estRows) {
                    if (TryReadSingleRangeChunkXmlFast(r1, c1, r2, c2, ct, out var chunk)) {
                        if (chunk != null) {
                            yield return chunk;
                        }

                        yield break;
                    }
                }

                if (estRows <= BufferedRangeStreamRowLimit
                    && TryReadBufferedRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, estRows, ct, out var bufferedChunks)) {
                    foreach (var chunk in bufferedChunks) {
                        yield return chunk;
                    }

                    yield break;
                }

                if (ShouldUseOrderedBufferedXmlStream(estRows, c1, c2)
                    && TryReadOrderedBufferedRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, estRows, ct, out var automaticChunks)) {
                    foreach (var chunk in automaticChunks) {
                        yield return chunk;
                    }

                    yield break;
                }

                if (RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                    foreach (var chunk in ReadRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, ct)) {
                        yield return chunk;
                    }

                    yield break;
                }
            }

            int dop = (decided == OfficeIMO.Excel.ExecutionMode.Parallel)
                ? (policy.MaxDegreeOfParallelism ?? System.Environment.ProcessorCount)
                : 1;
            if (dop < 1) dop = 1;

            int nextToYield = 0;
            int chunkIndex = 0;

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel
                && estRows <= BufferedRangeStreamRowLimit
                && CanUseRangeStreamXmlReader()) {
                if (chunkRows >= estRows) {
                    if (TryReadSingleRangeChunkXmlFast(r1, c1, r2, c2, ct, out var chunk)) {
                        if (chunk != null) {
                            yield return chunk;
                        }

                        yield break;
                    }
                }

                if (TryReadBufferedRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, estRows, ct, out var chunks)) {
                    foreach (var chunk in chunks) {
                        yield return chunk;
                    }

                    yield break;
                }

                foreach (var chunk in ReadBufferedRangeStreamFromFastRange(a1Range, r1, c1, chunkRows, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            if (estRows <= BufferedRangeStreamRowLimit) {
                foreach (var chunk in ReadBufferedRows(EnumerateWorksheetRows(ct), r1, c1, r2, c2, decided, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel
                && CanUseRangeStreamXmlReader()
                && RowsAreSortedWithinRangeXmlFast(r1, r2, ct)) {
                foreach (var chunk in ReadRangeStreamXmlFast(r1, c1, r2, c2, chunkRows, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData is null) yield break;

            if (estRows > chunkRows && !RowsAreSortedWithinRange(sheetData, r1, r2, ct)) {
                foreach (var chunk in ReadUnsortedRows(sheetData, r1, c1, r2, c2, decided, ct)) {
                    yield return chunk;
                }

                yield break;
            }

            if (decided != OfficeIMO.Excel.ExecutionMode.Parallel) {
                int currentWindow = -1;
                var sequentialRows = new List<Row>();

                foreach (var row in sheetData.Elements<Row>()) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }

                    int rIdx = checked((int)row.RowIndex!.Value);
                    if (rIdx < r1) continue;
                    if (rIdx > r2) break;

                    int window = GetWindowIndex(rIdx);
                    if (currentWindow >= 0 && window != currentWindow) {
                        yield return ConvertChunk(sequentialRows, currentWindow, r1, c1, r2, c2, ct);
                        sequentialRows.Clear();
                    }

                    currentWindow = window;
                    sequentialRows.Add(row);
                }

                if (sequentialRows.Count > 0) {
                    yield return ConvertChunk(sequentialRows, currentWindow, r1, c1, r2, c2, ct);
                }

                yield break;
            }

            int maxPendingChunks = dop;
            var pending = new Dictionary<int, Task<RangeChunk>>(maxPendingChunks);
            int activeWindow = -1;
            List<Row> bufferRows = new();

            foreach (var row in sheetData.Elements<Row>()) {
                if (canCancel) {
                    ct.ThrowIfCancellationRequested();
                }

                int rIdx = checked((int)row.RowIndex!.Value);
                if (rIdx < r1) continue;
                if (rIdx > r2) break;

                int window = GetWindowIndex(rIdx);
                if (activeWindow >= 0 && window != activeWindow) {
                    ScheduleChunk(bufferRows, activeWindow, r1, c1, r2, c2);
                    bufferRows = new List<Row>();

                    while (pending.Count >= maxPendingChunks) {
                        yield return CompleteChunk(nextToYield++);
                    }
                }

                activeWindow = window;
                bufferRows.Add(row);
            }

            if (bufferRows.Count > 0) {
                ScheduleChunk(bufferRows, activeWindow, r1, c1, r2, c2);
            }

            while (pending.Count > 0) {
                yield return CompleteChunk(nextToYield++);
            }

            void ScheduleChunk(List<Row> rows, int windowIndex, int rr1, int cc1, int rr2, int cc2) {
                var snapshot = rows.ToArray();
                int scheduledIndex = chunkIndex++;
                pending.Add(scheduledIndex, Task.Run(() => ConvertChunk(snapshot, windowIndex, rr1, cc1, rr2, cc2, ct), ct));
            }

            RangeChunk CompleteChunk(int index) {
                if (!pending.TryGetValue(index, out var task)) {
                    throw new InvalidOperationException($"Chunk {index} was not scheduled.");
                }

                pending.Remove(index);
                try {
                    return task.GetAwaiter().GetResult();
                } catch (AggregateException ex) when (ex.InnerExceptions.Count == 1) {
                    throw ex.InnerExceptions[0];
                }
            }

            RangeChunk ConvertChunk(IReadOnlyList<Row> rows, int index, int rr1, int cc1, int rr2, int cc2, CancellationToken token) {
                bool chunkCanCancel = token.CanBeCanceled;
                if (chunkCanCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int startRow = GetWindowStartRow(index);
                int endRow = Math.Min(startRow + chunkRows - 1, rr2);
                int height = endRow - startRow + 1;
                int width = cc2 - cc1 + 1;
                if (height <= 0 || width <= 0)
                    return new RangeChunk(startRow, 0, cc1, width, Array.Empty<object?[]>());

                var outRows = new object?[height][];
                for (int i = 0; i < height; i++) outRows[i] = new object?[width];

                foreach (var rowEl in rows) {
                    if (chunkCanCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    int rowIndex = checked((int)rowEl.RowIndex!.Value);
                    int rowOffset = rowIndex - startRow;
                    if ((uint)rowOffset >= (uint)height) continue;

                    foreach (var cell in rowEl.Elements<Cell>()) {
                        int c = A1.ParseColumnIndexFromCellReferenceFast(cell.CellReference?.Value);
                        if (c < cc1 || c > cc2) continue;
                        if (TryConvertCell(cell, out object? value)) {
                            outRows[rowOffset][c - cc1] = value ?? outRows[rowOffset][c - cc1];
                        }
                    }
                }

                return new RangeChunk(startRow, height, cc1, width, outRows);
            }

            int GetWindowIndex(int rowIndex) => (rowIndex - r1) / chunkRows;

            int GetWindowStartRow(int index) => r1 + (index * chunkRows);

            IEnumerable<RangeChunk> ReadBufferedRangeStreamFromFastRange(string range, int firstRow, int firstColumn, int rowsPerChunk, CancellationToken token) {
                object?[,] values = ReadRange(range, OfficeIMO.Excel.ExecutionMode.Sequential, token);
                int height = values.GetLength(0);
                int width = values.GetLength(1);
                bool matrixCanCancel = token.CanBeCanceled;

                for (int offset = 0; offset < height; offset += rowsPerChunk) {
                    if (matrixCanCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    int rowCount = Math.Min(rowsPerChunk, height - offset);
                    var outRows = new object?[rowCount][];
                    for (int r = 0; r < rowCount; r++) {
                        var rowValues = new object?[width];
                        for (int c = 0; c < width; c++) {
                            rowValues[c] = values[offset + r, c];
                        }

                        outRows[r] = rowValues;
                    }

                    yield return new RangeChunk(firstRow + offset, rowCount, firstColumn, width, outRows);
                }
            }

            IEnumerable<RangeChunk> ReadBufferedRows(IEnumerable<Row> sourceRows, int rr1, int cc1, int rr2, int cc2, OfficeIMO.Excel.ExecutionMode executionMode, CancellationToken token) {
                int windowCount = GetWindowIndex(rr2) + 1;
                var windows = new List<Row>?[windowCount];

                foreach (var row in sourceRows) {
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    int rowIndex = checked((int)row.RowIndex!.Value);
                    if (rowIndex < rr1) continue;
                    if (rowIndex > rr2) continue;

                    int window = GetWindowIndex(rowIndex);
                    if ((uint)window >= (uint)windows.Length) continue;
                    windows[window] ??= new List<Row>();
                    windows[window]!.Add(row);
                }

                if (executionMode != OfficeIMO.Excel.ExecutionMode.Parallel) {
                    for (int i = 0; i < windows.Length; i++) {
                        var rows = windows[i];
                        if (rows == null) continue;

                        yield return ConvertChunk(rows, i, rr1, cc1, rr2, cc2, token);
                    }

                    yield break;
                }

                int maxBufferedPendingChunks = dop;
                var bufferedPending = new Dictionary<int, Task<RangeChunk>>(maxBufferedPendingChunks);
                int bufferedScheduleIndex = 0;
                int bufferedNextToYield = 0;

                for (int i = 0; i < windows.Length; i++) {
                    var rows = windows[i];
                    if (rows == null) continue;

                    int scheduledIndex = bufferedScheduleIndex++;
                    Row[] snapshot = rows.ToArray();
                    int windowIndex = i;
                    bufferedPending.Add(scheduledIndex, Task.Run(() => ConvertChunk(snapshot, windowIndex, rr1, cc1, rr2, cc2, token), token));

                    while (bufferedPending.Count >= maxBufferedPendingChunks) {
                        yield return CompleteBufferedChunk(bufferedNextToYield++);
                    }
                }

                while (bufferedPending.Count > 0) {
                    yield return CompleteBufferedChunk(bufferedNextToYield++);
                }

                RangeChunk CompleteBufferedChunk(int index) {
                    if (!bufferedPending.TryGetValue(index, out var task)) {
                        throw new InvalidOperationException($"Chunk {index} was not scheduled.");
                    }

                    bufferedPending.Remove(index);
                    try {
                        return task.GetAwaiter().GetResult();
                    } catch (AggregateException ex) when (ex.InnerExceptions.Count == 1) {
                        throw ex.InnerExceptions[0];
                    }
                }
            }

            IEnumerable<RangeChunk> ReadUnsortedRows(SheetData data, int rr1, int cc1, int rr2, int cc2, OfficeIMO.Excel.ExecutionMode executionMode, CancellationToken token) {
                var windows = new SortedDictionary<int, List<Row>>();
                foreach (var row in data.Elements<Row>()) {
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    int rowIndex = checked((int)row.RowIndex!.Value);
                    if (rowIndex < rr1) continue;
                    if (rowIndex > rr2) continue;

                    int window = GetWindowIndex(rowIndex);
                    if (!windows.TryGetValue(window, out var windowRows)) {
                        windowRows = new List<Row>();
                        windows[window] = windowRows;
                    }

                    windowRows.Add(row);
                }

                if (executionMode != OfficeIMO.Excel.ExecutionMode.Parallel) {
                    foreach (var window in windows) {
                        yield return ConvertChunk(window.Value, window.Key, rr1, cc1, rr2, cc2, token);
                    }

                    yield break;
                }

                int maxUnsortedPendingChunks = dop;
                var unsortedPending = new Dictionary<int, Task<RangeChunk>>(maxUnsortedPendingChunks);
                int unsortedScheduleIndex = 0;
                int unsortedNextToYield = 0;

                foreach (var window in windows) {
                    int scheduledIndex = unsortedScheduleIndex++;
                    Row[] snapshot = window.Value.ToArray();
                    int windowIndex = window.Key;
                    unsortedPending.Add(scheduledIndex, Task.Run(() => ConvertChunk(snapshot, windowIndex, rr1, cc1, rr2, cc2, token), token));

                    while (unsortedPending.Count >= maxUnsortedPendingChunks) {
                        yield return CompleteUnsortedChunk(unsortedNextToYield++);
                    }
                }

                while (unsortedPending.Count > 0) {
                    yield return CompleteUnsortedChunk(unsortedNextToYield++);
                }

                RangeChunk CompleteUnsortedChunk(int index) {
                    if (!unsortedPending.TryGetValue(index, out var task)) {
                        throw new InvalidOperationException($"Chunk {index} was not scheduled.");
                    }

                    unsortedPending.Remove(index);
                    try {
                        return task.GetAwaiter().GetResult();
                    } catch (AggregateException ex) when (ex.InnerExceptions.Count == 1) {
                        throw ex.InnerExceptions[0];
                    }
                }
            }
        }

        private static bool ShouldUseOrderedBufferedXmlStream(int estimatedRows, int firstColumn, int lastColumn) {
            int width = lastColumn - firstColumn + 1;
            return width > 0
                && estimatedRows > 0
                && ((long)estimatedRows * width) <= OrderedBufferedRangeStreamCellLimit;
        }

        private bool CanUseAutomaticXmlStreamFastPath(bool automaticDecision, OfficeIMO.Excel.ExecutionMode decided) {
            return automaticDecision
                && decided != OfficeIMO.Excel.ExecutionMode.Parallel
                && CanAttemptRangeStreamXmlReader();
        }

        private bool CanUseRangeStreamXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && CanStreamWorksheetPart();
        }

        private bool CanAttemptRangeStreamXmlReader() {
            return (_opt.CellValueConverter != null || _opt.Culture == System.Globalization.CultureInfo.InvariantCulture)
                && _canStreamWorksheetPart
                && _hasWorksheetPartStreamContent != false;
        }

        private bool TryReadSingleRangeChunkXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            CancellationToken ct,
            out RangeChunk? chunk) {
            chunk = null;
            int width = c2 - c1 + 1;
            object?[][]? rows = null;

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                if (!TryPrepareWorksheetStream(stream)) {
                    _hasWorksheetPartStreamContent = false;
                    return false;
                }

                _hasWorksheetPartStreamContent = true;
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                int requestedRowCount = r2 - r1 + 1;
                var seenRows = CreateCompletedRowTracker(requestedRowCount);

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
                            if (rowIndex > r2 && seenRows.AllRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        if (rows == null) {
                            rows = new object?[requestedRowCount][];
                            for (int row = 0; row < rows.Length; row++) {
                                rows[row] = new object?[width];
                            }
                        }

                        ReadXmlRowIntoChunk(reader, rows, rowIndex, r1, c1, c2, ct);
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
                            if (rowIndex > r2 && seenRows.AllRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        if (rows == null) {
                            rows = new object?[requestedRowCount][];
                            for (int row = 0; row < rows.Length; row++) {
                                rows[row] = new object?[width];
                            }
                        }

                        ReadXmlRowIntoChunk(reader, rows, rowIndex, r1, c1, c2, CancellationToken.None);
                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                }

                if (rows != null) {
                    chunk = new RangeChunk(r1, rows.Length, c1, width, rows);
                }

                return true;
            } catch (XmlException) {
                chunk = null;
                return false;
            } catch (IOException) {
                chunk = null;
                return false;
            } catch (UnauthorizedAccessException) {
                chunk = null;
                return false;
            } catch (ObjectDisposedException) {
                chunk = null;
                return false;
            }
        }

        private IEnumerable<RangeChunk> ReadRangeStreamXmlFast(int r1, int c1, int r2, int c2, int chunkRows, CancellationToken ct) {
            using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
            RewindWorksheetStream(stream);
            using var reader = OpenWorksheetXmlReader(stream);
            bool canCancel = ct.CanBeCanceled;
            int width = c2 - c1 + 1;
            int currentWindow = -1;
            int currentStartRow = 0;
            object?[][]? currentRows = null;
            int nextRowIndex = 1;
            int requestedRowCount = r2 - r1 + 1;
            var seenRows = CreateCompletedRowTracker(requestedRowCount);

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

                int window = (rowIndex - r1) / chunkRows;
                if (currentRows != null && window != currentWindow) {
                    yield return new RangeChunk(currentStartRow, currentRows.Length, c1, width, currentRows);
                    currentRows = null;
                }

                if (currentRows == null) {
                    currentWindow = window;
                    currentStartRow = r1 + (window * chunkRows);
                    int rowCount = Math.Min(chunkRows, r2 - currentStartRow + 1);
                    currentRows = new object?[rowCount][];
                    for (int i = 0; i < rowCount; i++) {
                        currentRows[i] = new object?[width];
                    }
                }

                ReadXmlRowIntoChunk(reader, currentRows, rowIndex, currentStartRow, c1, c2, ct);
                seenRows.MarkSeen(rowIndex - r1);
                if (seenRows.AllRowsSeen) {
                    break;
                }
            }

            if (currentRows != null) {
                yield return new RangeChunk(currentStartRow, currentRows.Length, c1, width, currentRows);
            }
        }

        private bool TryReadOrderedBufferedRangeStreamXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            int estimatedRows,
            CancellationToken ct,
            out RangeChunk[] chunks) {
            chunks = Array.Empty<RangeChunk>();
            int width = c2 - c1 + 1;
            var chunkMap = new Dictionary<int, RangeChunk>(Math.Max(1, Math.Min(((estimatedRows - 1) / chunkRows) + 1, 256)));

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(estimatedRows);
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

                    int window = (rowIndex - r1) / chunkRows;
                    if (!chunkMap.TryGetValue(window, out var chunk)) {
                        int startRow = r1 + (window * chunkRows);
                        int rowCount = Math.Min(chunkRows, r2 - startRow + 1);
                        var rows = new object?[rowCount][];
                        for (int row = 0; row < rowCount; row++) {
                            rows[row] = new object?[width];
                        }

                        chunk = new RangeChunk(startRow, rowCount, c1, width, rows);
                        chunkMap.Add(window, chunk);
                    }

                    ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, ct);
                    seenRows.MarkSeen(rowIndex - r1);
                    if (seenRows.AllRowsSeen) {
                        break;
                    }
                }

                if (chunkMap.Count == 0) {
                    return true;
                }

                int index = 0;
                chunks = new RangeChunk[chunkMap.Count];
                int[] windows = new int[chunkMap.Count];
                chunkMap.Keys.CopyTo(windows, 0);
                Array.Sort(windows);
                foreach (int window in windows) {
                    chunks[index++] = chunkMap[window];
                }

                return true;
            } catch (XmlException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (IOException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (UnauthorizedAccessException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (ObjectDisposedException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            }
        }

        private bool TryReadBufferedRangeStreamXmlFast(
            int r1,
            int c1,
            int r2,
            int c2,
            int chunkRows,
            int estimatedRows,
            CancellationToken ct,
            out RangeChunk[] chunks) {
            int width = c2 - c1 + 1;
            int windowCount = ((estimatedRows - 1) / chunkRows) + 1;
            chunks = new RangeChunk[windowCount];
            for (int window = 0; window < chunks.Length; window++) {
                int startRow = r1 + (window * chunkRows);
                int rowCount = Math.Min(chunkRows, r2 - startRow + 1);
                var rows = new object?[rowCount][];
                for (int row = 0; row < rowCount; row++) {
                    rows[row] = new object?[width];
                }

                chunks[window] = new RangeChunk(startRow, rowCount, c1, width, rows);
            }

            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = ct.CanBeCanceled;
                int nextRowIndex = 1;
                var seenRows = CreateCompletedRowTracker(estimatedRows);
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
                            if (rowIndex > r2 && seenRows.AllRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        int window = (rowIndex - r1) / chunkRows;
                        if ((uint)window >= (uint)chunks.Length) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        var chunk = chunks[window];
                        ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, ct);
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
                            if (rowIndex > r2 && seenRows.AllRowsSeen) {
                                break;
                            }

                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        int window = (rowIndex - r1) / chunkRows;
                        if ((uint)window >= (uint)chunks.Length) {
                            SkipXmlElement(reader, "row");
                            continue;
                        }

                        var chunk = chunks[window];
                        ReadXmlRowIntoChunk(reader, chunk.Rows, rowIndex, chunk.StartRow, c1, c2, CancellationToken.None);
                        seenRows.MarkSeen(rowIndex - r1);
                        if (seenRows.AllRowsSeen) {
                            break;
                        }
                    }
                }

                return true;
            } catch (XmlException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (IOException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (UnauthorizedAccessException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            } catch (ObjectDisposedException) {
                chunks = Array.Empty<RangeChunk>();
                return false;
            }
        }

        private void ReadXmlRowIntoChunk(XmlReader rowReader, object?[][] rows, int rowIndex, int startRow, int c1, int c2, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int rowOffset = rowIndex - startRow;
            if ((uint)rowOffset >= (uint)rows.Length) {
                return;
            }

            object?[] rowValues = rows[rowOffset];
            if (rowValues.Length == 8) {
                ReadXmlRowIntoChunk8(rowReader, rowValues, c1, c2, ct);
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            bool canTrackColumns = rowValues.Length <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(rowValues.Length) : 0UL;
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

                int columnOffset = columnIndex - c1;
                if ((uint)columnOffset >= (uint)rowValues.Length) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedFullWidthExit && columnIndex != nextExpectedColumn) {
                    canUseOrderedFullWidthExit = false;
                    int orderedSeen = nextExpectedColumn - c1;
                    seenColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

                rowValues[columnOffset] = ReadXmlCellValue(rowReader);
                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                }

                if (canUseOrderedFullWidthExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }

                if (canTrackColumns && !canUseOrderedFullWidthExit && MarkRequestedColumnSeen(columnOffset, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private void ReadXmlRowIntoChunk8(XmlReader rowReader, object?[] rowValues, int c1, int c2, CancellationToken ct) {
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
                    canUseOrderedFullWidthExit = false;
                }

                rowValues[columnOffset] = ReadXmlCellValue(rowReader);
                seenColumns |= 1UL << columnOffset;

                if (canUseOrderedFullWidthExit) {
                    nextExpectedColumn++;
                    if (columnIndex >= c2) {
                        SkipXmlElementContent(rowReader, depth, "row");
                        return;
                    }
                } else if (seenColumns == 0xFFUL) {
                    SkipXmlElementContent(rowReader, depth, "row");
                    return;
                }
            }
        }

        private bool RowsAreSortedWithinRangeXmlFast(int firstRow, int lastRow, CancellationToken token) {
            try {
                using var stream = _wsPart.GetStream(FileMode.Open, FileAccess.Read);
                RewindWorksheetStream(stream);
                using var reader = OpenWorksheetXmlReader(stream);
                bool canCancel = token.CanBeCanceled;
                bool hasPrevious = false;
                bool sawRowAfterRange = false;
                int previous = 0;
                int nextRowIndex = 1;
                int rowCount = lastRow - firstRow + 1;
                var seenRows = CreateCompletedRowTracker(rowCount);

                while (reader.Read()) {
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row") {
                        continue;
                    }

                    int rowIndex = ParsePositiveIntAttribute(reader.GetAttribute("r"));
                    if (rowIndex <= 0) {
                        rowIndex = nextRowIndex;
                    }

                    nextRowIndex = rowIndex + 1;
                    if (rowIndex < firstRow) {
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (rowIndex > lastRow) {
                        if (seenRows.AllRowsSeen) {
                            return true;
                        }

                        sawRowAfterRange = true;
                        SkipXmlElement(reader, "row");
                        continue;
                    }

                    if (sawRowAfterRange) {
                        return false;
                    }

                    if (hasPrevious && rowIndex <= previous) {
                        return false;
                    }

                    previous = rowIndex;
                    hasPrevious = true;
                    seenRows.MarkSeen(rowIndex - firstRow);
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

        private static bool RowsAreSortedWithinRange(SheetData data, int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;
            int rowCount = lastRow - firstRow + 1;
            var seenRows = CreateCompletedRowTracker(rowCount);

            foreach (var row in data.Elements<Row>()) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    if (seenRows.AllRowsSeen) {
                        return true;
                    }

                    sawRowAfterRange = true;
                    continue;
                }
                if (sawRowAfterRange) {
                    return false;
                }

                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
                seenRows.MarkSeen(rowIndex - firstRow);
            }

            return true;
        }

        private bool RowsAreSortedWithinRange(int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            bool sawRowAfterRange = false;
            int previous = 0;
            int rowCount = lastRow - firstRow + 1;
            var seenRows = CreateCompletedRowTracker(rowCount);

            foreach (var row in EnumerateWorksheetRows(token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) {
                    if (seenRows.AllRowsSeen) {
                        return true;
                    }

                    sawRowAfterRange = true;
                    continue;
                }
                if (sawRowAfterRange) {
                    return false;
                }

                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
                seenRows.MarkSeen(rowIndex - firstRow);
            }

            return true;
        }

        /// <summary>
        /// Represents a rectangular block of rows produced during streaming.
        /// </summary>
        public sealed class RangeChunk {
            /// <summary>First row index (1-based) covered by this chunk.</summary>
            public int StartRow { get; }
            /// <summary>Number of rows in this chunk.</summary>
            public int RowCount { get; }
            /// <summary>First column index (1-based) covered by this chunk.</summary>
            public int StartCol { get; }
            /// <summary>Number of columns in this chunk.</summary>
            public int ColCount { get; }
            /// <summary>Row-major values array. Size is <see cref="RowCount"/> x <see cref="ColCount"/>.</summary>
            public object?[][] Rows { get; }

            internal RangeChunk(int startRow, int rowCount, int startCol, int colCount, object?[][] rows) {
                StartRow = startRow; RowCount = rowCount; StartCol = startCol; ColCount = colCount; Rows = rows;
            }
        }
    }
}
