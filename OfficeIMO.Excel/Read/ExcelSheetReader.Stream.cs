using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Streaming APIs for large ranges.
    /// </summary>
    public sealed partial class ExcelSheetReader
    {
        /// <summary>
        /// Lazily reads a rectangular A1 range as ordered row chunks. DOM traversal is single-threaded;
        /// per-chunk value conversion is offloaded in parallel based on Execution policy.
        /// </summary>
        public IEnumerable<RangeChunk> ReadRangeStream(string a1Range, int chunkRows = 1024, OfficeIMO.Excel.ExecutionMode? mode = null, CancellationToken ct = default)
        {
            (int r1, int c1, int r2, int c2) = A1.ParseRange(a1Range);
            if (r1 > r2 || c1 > c2) yield break;

            int estRows = Math.Max(0, r2 - r1 + 1);
            var policy = _opt.Execution;
            var decided = mode ?? policy.Mode;
            if (decided == OfficeIMO.Excel.ExecutionMode.Automatic)
                decided = policy.Decide("ReadRangeStream", estRows);

            int dop = (decided == OfficeIMO.Excel.ExecutionMode.Parallel)
                ? (policy.MaxDegreeOfParallelism ?? System.Environment.ProcessorCount)
                : 1;
            if (dop < 1) dop = 1;

            using var semaphore = new SemaphoreSlim(dop, dop);
            var tasks = new List<Task>();
            var results = new ConcurrentDictionary<int, RangeChunk>(); // chunkIndex -> chunk
            int nextToYield = 0;
            int chunkIndex = 0;

            var sheetData = _wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData is null) yield break;

            List<Row> bufferRows = new(chunkRows);

            foreach (var row in sheetData.Elements<Row>())
            {
                if (ct.IsCancellationRequested) yield break;

                int rIdx = checked((int)row.RowIndex!.Value);
                if (rIdx < r1) continue;
                if (rIdx > r2) break;

                bufferRows.Add(row);
                if (bufferRows.Count >= chunkRows)
                {
                    ScheduleChunk(bufferRows, chunkIndex++, r1, c1, r2, c2);
                    bufferRows = new List<Row>(chunkRows);
                }
            }

            if (bufferRows.Count > 0)
                ScheduleChunk(bufferRows, chunkIndex++, r1, c1, r2, c2);

            for (int i = 0; i < chunkIndex; i++)
            {
                RangeChunk? readyChunk;
                while (!results.TryRemove(nextToYield, out readyChunk))
                {
                    Thread.SpinWait(200);
                    Thread.Yield();
                }
                yield return readyChunk!;
                nextToYield++;
            }

            void ScheduleChunk(List<Row> rows, int index, int rr1, int cc1, int rr2, int cc2)
            {
                var snapshot = rows.ToArray();
                tasks.Add(Task.Run(async () =>
                {
                    await semaphore.WaitAsync(ct).ConfigureAwait(false);
                    try
                    {
                        var chunk = ConvertChunk(snapshot, index, rr1, cc1, rr2, cc2, ct);
                        results[index] = chunk;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                }, ct));
            }

            RangeChunk ConvertChunk(Row[] rows, int index, int rr1, int cc1, int rr2, int cc2, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();

                int startRow = rows.Length > 0 ? (int)rows[0].RowIndex!.Value : rr1;
                startRow = Math.Max(startRow, rr1);

                int endRow = rows.Length > 0 ? (int)rows[rows.Length - 1].RowIndex!.Value : startRow;
                endRow = Math.Min(endRow, rr2);

                int height = endRow - startRow + 1;
                int width = cc2 - cc1 + 1;
                if (height <= 0 || width <= 0)
                    return new RangeChunk(startRow, 0, cc1, width, Array.Empty<object?[]>());

                var rowMap = new Dictionary<int, Row>(rows.Length);
                foreach (var r in rows)
                {
                    int ridx = (int)r.RowIndex!.Value;
                    if (ridx >= rr1 && ridx <= rr2) rowMap[ridx] = r;
                }

                var outRows = new object?[height][];
                for (int i = 0; i < height; i++) outRows[i] = new object?[width];

                for (int i = 0; i < height; i++)
                {
                    token.ThrowIfCancellationRequested();
                    int absoluteRow = startRow + i;
                    if (!rowMap.TryGetValue(absoluteRow, out var rowEl)) continue;

                    foreach (var cell in rowEl.Elements<Cell>())
                    {
                        if (cell.CellReference?.Value is null) continue;
                        var (r, c) = A1.ParseCellRef(cell.CellReference.Value);
                        if (c < cc1 || c > cc2) continue;
                        var val = ConvertCell(cell);
                        outRows[i][c - cc1] = val ?? outRows[i][c - cc1];
                    }
                }

                return new RangeChunk(startRow, height, cc1, width, outRows);
            }
        }

        /// <summary>
        /// Represents a rectangular block of rows produced during streaming.
        /// </summary>
        public sealed class RangeChunk
        {
            public int StartRow { get; }
            public int RowCount { get; }
            public int StartCol { get; }
            public int ColCount { get; }
            public object?[][] Rows { get; }

            internal RangeChunk(int startRow, int rowCount, int startCol, int colCount, object?[][] rows)
            {
                StartRow = startRow; RowCount = rowCount; StartCol = startCol; ColCount = colCount; Rows = rows;
            }
        }
    }
}
