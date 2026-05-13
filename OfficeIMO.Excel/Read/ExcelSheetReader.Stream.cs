using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private const int BufferedRangeStreamRowLimit = 4_096;

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
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic)
                decided = policy.Decide("ReadRangeStream", estRows);

            int dop = (decided == OfficeIMO.Excel.ExecutionMode.Parallel)
                ? (policy.MaxDegreeOfParallelism ?? System.Environment.ProcessorCount)
                : 1;
            if (dop < 1) dop = 1;

            int nextToYield = 0;
            int chunkIndex = 0;

            if (estRows <= BufferedRangeStreamRowLimit) {
                foreach (var chunk in ReadBufferedRows(EnumerateWorksheetRows(ct), r1, c1, r2, c2, decided, ct)) {
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

            int maxPendingChunks = Math.Max(dop, dop * 2);
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

            IEnumerable<RangeChunk> ReadBufferedRows(IEnumerable<Row> sourceRows, int rr1, int cc1, int rr2, int cc2, OfficeIMO.Excel.ExecutionMode executionMode, CancellationToken token) {
                int windowCount = GetWindowIndex(rr2) + 1;
                var windows = new List<Row>?[windowCount];

                foreach (var row in sourceRows) {
                    if (canCancel) {
                        token.ThrowIfCancellationRequested();
                    }

                    int rowIndex = checked((int)row.RowIndex!.Value);
                    if (rowIndex < rr1) continue;
                    if (rowIndex > rr2) break;

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

                int maxBufferedPendingChunks = Math.Max(dop, dop * 2);
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
                    if (rowIndex > rr2) break;

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

                int maxUnsortedPendingChunks = Math.Max(dop, dop * 2);
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

        private static bool RowsAreSortedWithinRange(SheetData data, int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            int previous = 0;

            foreach (var row in data.Elements<Row>()) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) break;
                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
            }

            return true;
        }

        private bool RowsAreSortedWithinRange(int firstRow, int lastRow, CancellationToken token) {
            bool canCancel = token.CanBeCanceled;
            bool hasPrevious = false;
            int previous = 0;

            foreach (var row in EnumerateWorksheetRows(token)) {
                if (canCancel) {
                    token.ThrowIfCancellationRequested();
                }

                int rowIndex = checked((int)row.RowIndex!.Value);
                if (rowIndex < firstRow) continue;
                if (rowIndex > lastRow) break;
                if (hasPrevious && rowIndex <= previous) {
                    return false;
                }

                previous = rowIndex;
                hasPrevious = true;
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
