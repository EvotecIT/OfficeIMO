#nullable enable

using System.Buffers;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const int MinimumParallelIndexBytes = 4 * 1024 * 1024;
            private const int MaximumParallelIndexHelpers = 3;
            private ParallelRowIndex? _parallelIndex;
            // Helpers share package cursors with the sequential indexer, so they must not load parts.
            private bool _isParallelIndexHelper;
            private bool _parallelIndexAttempted;
            // The complete part length when indexing overlaps inflation; zero uses the buffered length.
            private int _expectedLength;

            /// <summary>
            /// Starts helper workers that index later byte chunks of a uniform worksheet into
            /// private metadata. The sequential indexer remains authoritative: it adopts a chunk
            /// only when that chunk's first row starts exactly where its own parse arrived, so a
            /// helper can only repeat work the sequential pass would do with the same row shape.
            /// </summary>
            private void TryStartParallelIndex(int firstRow, int lastRow, int position) {
                if (_parallelIndexAttempted) return;
                _parallelIndexAttempted = true;
                int expected = _expectedLength > 0 ? _expectedLength : _length;
                if (expected < MinimumParallelIndexBytes || UsesCellWindow || _unknownRowExtent
                    || _worksheetPrefixLength != 0 || _formulaStarts != null || _repeatedRowSuffixStart < 0
                    || !_options.Execution.AllowsBackgroundProcessing) return;
                int helpers = Math.Min(MaximumParallelIndexHelpers, Environment.ProcessorCount - 2);
                int? maximumDegree = _options.Execution.MaxDegreeOfParallelism;
                if (maximumDegree.HasValue) helpers = Math.Min(helpers, maximumDegree.Value - 1);
                if (helpers <= 0) return;
                int chunkBytes = Math.Max(512 * 1024, Math.Min(4 * 1024 * 1024, expected / (helpers * 8)));
                int averageRowBytes = Math.Max(32, (position - Math.Max(0, _sheetDataContentStart)) / Math.Max(1, _rowCount));
                _parallelIndex = new ParallelRowIndex(this, expected, chunkBytes, firstChunkStart: position,
                    averageRowBytes, firstRow, lastRow, helpers, _options.CancellationToken);
            }

            /// <summary>Adopts the helper chunk that starts at the current row boundary, if it qualifies.</summary>
            private bool TryStitchParallelChunk(ref int position, ref int previousRow, CancellationToken ct) {
                ParallelRowIndex index = _parallelIndex!;
                ParallelRowChunk? chunk = index.TakeChunkAt(position, ct);
                if (chunk == null) return false;
                try {
                    if (chunk.RowCount == 0 || chunk.Start != position || chunk.RowIndexes![0] <= previousRow) return false;
                    while (_length < chunk.End) {
                        if (!TryAdvanceInflation(ct)) return false;
                    }

                    int rows = chunk.RowCount;
                    int cells = checked(rows * _fieldCount);
                    EnsureRowCapacity(_rowCount + rows);
                    int rowOffset = GetQualificationRowOffset(_rowCount);
                    Array.Copy(chunk.RowIndexes!, 0, _rowIndexes!, _rowCount, rows);
                    Array.Copy(chunk.ValueStarts!, 0, _valueStarts!, rowOffset, cells);
                    Array.Copy(chunk.ValueLengths!, 0, _valueLengths!, rowOffset, cells);
                    Array.Copy(chunk.CellKinds!, 0, _cellKinds!, rowOffset, cells);
                    if (chunk.DecodedTags != null) {
                        EnsureDecodedValues();
                        ClearDecodedValuesThrough(rowOffset);
                        Array.Copy(chunk.DecodedValues!, 0, _decodedValues!, rowOffset, cells);
                        Array.Copy(chunk.DecodedTags, 0, _decodedTags!, rowOffset, cells);
                        _decodedThroughCell = rowOffset + cells;
                    }
                    if (_formulaStarts != null) {
                        Array.Clear(_formulaStarts, rowOffset, cells);
                        _formulaLengths.AsSpan(rowOffset, cells).Fill(-1);
                    }

                    if (chunk.MinimumCellColumn < _minimumCellColumn) _minimumCellColumn = chunk.MinimumCellColumn;
                    if (chunk.MaximumCellRow > 0) {
                        if (_minimumCellRow == int.MaxValue) _minimumCellRow = chunk.MinimumCellRow;
                        _maximumCellRow = chunk.MaximumCellRow;
                        if (chunk.MaximumCellColumn > _maximumCellColumn) _maximumCellColumn = chunk.MaximumCellColumn;
                    }

                    previousRow = chunk.RowIndexes[rows - 1];
                    _rowCount += rows;
                    position = chunk.End;
                    return true;
                } finally {
                    chunk.ReturnStorage();
                }
            }

            private void StopParallelIndex() {
                ParallelRowIndex? index = _parallelIndex;
                if (index == null) return;
                _parallelIndex = null;
                index.Dispose();
                // Rows indexed sequentially after the last adopted chunk keep ordinary parsing.
                if (!UsesCellWindow) ClearDecodedValuesThrough(checked(_rowCount * _fieldCount));
            }

            /// <summary>Indexes one helper chunk with the sequential path's row and cell routines.</summary>
            private void IndexParallelChunk(ParallelRowIndex index, ParallelRowChunk chunk, int chunkStart, int chunkEnd, CancellationToken ct) {
                int available = index.WaitForBytes(0, ct);
                while (available < chunkStart + 16 && !index.IsComplete(available)) {
                    int next = index.WaitForBytes(available, ct);
                    if (next <= available) return;
                    available = next;
                }

                int position = -1;
                int searchFrom = chunkStart;
                while (true) {
                    int searchEnd = Math.Min(available, chunkEnd);
                    if (searchFrom < searchEnd) {
                        int relative = _buffer!.AsSpan(searchFrom, searchEnd - searchFrom).IndexOf("<row "u8);
                        if (relative >= 0) {
                            position = searchFrom + relative;
                            break;
                        }
                        searchFrom = Math.Max(searchFrom, searchEnd - 4);
                    }
                    if (available >= chunkEnd || index.IsComplete(available)) return;
                    int next = index.WaitForBytes(available, ct);
                    if (next <= available) return;
                    available = next;
                }

                _length = available;
                _minimumCellRow = int.MaxValue;
                _maximumCellRow = 0;
                _minimumCellColumn = int.MaxValue;
                _maximumCellColumn = 0;
                chunk.Start = position;
                int rows = 0;
                int previousRow = 0;
                chunk.EnsureCapacity(1, _fieldCount, index.DecodeValues);
                while (position < chunkEnd && !index.StopRequested) {
                    if (position >= _length) {
                        int next = index.WaitForBytes(_length, ct);
                        if (next <= _length) break;
                        _length = next;
                    }

                    int rowStart = position;
                    if (!TryReadRepeatedRowStartTag(ref position, out Utf8Tag tag, out int rowIndex)
                        || rowIndex <= previousRow || rowIndex < index.FirstRow || rowIndex > index.LastRow) {
                        position = rowStart;
                        break;
                    }

                    chunk.EnsureCapacity(rows + 1, _fieldCount, index.DecodeValues);
                    _rowIndexes = chunk.RowIndexes;
                    _valueStarts = chunk.ValueStarts;
                    _valueLengths = chunk.ValueLengths;
                    _cellKinds = chunk.CellKinds;
                    int rowOffset = rows * _fieldCount;
                    InitializeMetadataRow(rowOffset);
                    if (!tag.IsEmpty) {
                        bool indexed;
                        try {
                            indexed = TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
                        } catch (Exception) {
                            // The sequential indexer reparses this row and reports its own failure.
                            indexed = false;
                        }
                        if (!indexed) {
                            position = rowStart;
                            break;
                        }
                    }

                    if (index.DecodeValues) DecodeRowValues(rowOffset, chunk.DecodedValues!, chunk.DecodedTags!);
                    chunk.RowIndexes![rows++] = rowIndex;
                    previousRow = rowIndex;
                }

                chunk.End = position;
                chunk.RowCount = rows;
                chunk.MinimumCellRow = _minimumCellRow;
                chunk.MaximumCellRow = _maximumCellRow;
                chunk.MinimumCellColumn = _minimumCellColumn;
                chunk.MaximumCellColumn = _maximumCellColumn;
            }

            private sealed class ParallelIndexRowDeferredException : Exception { }

            private sealed class ParallelRowChunk {
                internal const int Unclaimed = 0;
                internal const int Running = 1;
                internal const int Completed = 2;
                internal const int OwnedBySequentialIndexer = 3;

                internal int State;
                internal int Start = -1;
                internal int End;
                internal int RowCount;
                internal int MinimumCellRow = int.MaxValue;
                internal int MaximumCellRow;
                internal int MinimumCellColumn = int.MaxValue;
                internal int MaximumCellColumn;
                internal int EstimatedRows;
                internal int[]? RowIndexes;
                internal int[]? ValueStarts;
                internal int[]? ValueLengths;
                internal byte[]? CellKinds;
                internal long[]? DecodedValues;
                internal byte[]? DecodedTags;

                internal void EnsureCapacity(int rows, int fieldCount, bool decodeValues) {
                    int current = RowIndexes?.Length ?? 0;
                    if (rows <= current && rows * fieldCount <= ValueStarts!.Length) return;
                    int nextRows = Math.Max(rows, Math.Max(EstimatedRows, current * 2));
                    int cells = checked(nextRows * fieldCount);
                    int used = RowIndexes == null ? 0 : RowCountForCopy(rows);
                    RowIndexes = Grow(RowIndexes, nextRows, used);
                    ValueStarts = Grow(ValueStarts, cells, used * fieldCount);
                    ValueLengths = Grow(ValueLengths, cells, used * fieldCount);
                    byte[] kinds = ArrayPool<byte>.Shared.Rent(cells);
                    if (CellKinds != null) {
                        Array.Copy(CellKinds, kinds, used * fieldCount);
                        ArrayPool<byte>.Shared.Return(CellKinds);
                    }
                    CellKinds = kinds;
                    if (decodeValues) {
                        long[] values = ArrayPool<long>.Shared.Rent(cells);
                        byte[] tags = ArrayPool<byte>.Shared.Rent(cells);
                        if (DecodedTags != null) {
                            Array.Copy(DecodedValues!, values, used * fieldCount);
                            Array.Copy(DecodedTags, tags, used * fieldCount);
                            ArrayPool<long>.Shared.Return(DecodedValues!);
                            ArrayPool<byte>.Shared.Return(DecodedTags);
                        }
                        DecodedValues = values;
                        DecodedTags = tags;
                    }
                }

                // Rows before the requested one are complete; the requested row is not yet written.
                private static int RowCountForCopy(int requestedRows) => requestedRows - 1;

                private static int[] Grow(int[]? values, int capacity, int count) {
                    int[] next = ArrayPool<int>.Shared.Rent(capacity);
                    if (values != null) {
                        Array.Copy(values, next, count);
                        ArrayPool<int>.Shared.Return(values);
                    }
                    return next;
                }

                internal void ReturnStorage() {
                    if (RowIndexes != null) ArrayPool<int>.Shared.Return(RowIndexes);
                    if (ValueStarts != null) ArrayPool<int>.Shared.Return(ValueStarts);
                    if (ValueLengths != null) ArrayPool<int>.Shared.Return(ValueLengths);
                    if (CellKinds != null) ArrayPool<byte>.Shared.Return(CellKinds);
                    if (DecodedValues != null) ArrayPool<long>.Shared.Return(DecodedValues);
                    if (DecodedTags != null) ArrayPool<byte>.Shared.Return(DecodedTags);
                    RowIndexes = ValueStarts = ValueLengths = null;
                    CellKinds = DecodedTags = null;
                    DecodedValues = null;
                }
            }

            private sealed class ParallelRowIndex : IDisposable {
                private readonly object _gate = new object();
                private readonly WorksheetInflationProgress? _inflation;
                private readonly int _bufferedLength;
                private readonly int _expectedLength;
                private readonly int _chunkBytes;
                private readonly int _firstChunkStart;
                private readonly ParallelRowChunk[] _chunks;
                private readonly Task[] _helpers;
                private readonly CancellationTokenSource _stop;
                private int _nextClaim;
                private int _nextChunk;
                private volatile bool _stopRequested;

                internal ParallelRowIndex(ExcelUtf8RangeRowSource owner, int expectedLength, int chunkBytes,
                    int firstChunkStart, int averageRowBytes, int firstRow, int lastRow, int helpers,
                    CancellationToken ct) {
                    _inflation = owner._inflation;
                    _bufferedLength = owner._length;
                    _expectedLength = expectedLength;
                    _chunkBytes = chunkBytes;
                    // Chunk zero begins at the current row and always belongs to the sequential indexer.
                    _firstChunkStart = firstChunkStart;
                    FirstRow = firstRow;
                    LastRow = lastRow;
                    int remaining = Math.Max(0, expectedLength - firstChunkStart);
                    int count = Math.Max(1, (remaining + chunkBytes - 1) / chunkBytes);
                    _chunks = new ParallelRowChunk[count];
                    int estimatedRows = Math.Max(64, chunkBytes / averageRowBytes + chunkBytes / (averageRowBytes * 4));
                    for (int i = 0; i < count; i++) _chunks[i] = new ParallelRowChunk { EstimatedRows = estimatedRows };
                    _chunks[0].State = ParallelRowChunk.OwnedBySequentialIndexer;
                    _nextChunk = 1;
                    _stop = CancellationTokenSource.CreateLinkedTokenSource(ct);
                    _helpers = new Task[helpers];
                    for (int i = 0; i < helpers; i++) {
                        ExcelUtf8RangeRowSource shadow = owner.CreateParallelIndexShadow();
                        _helpers[i] = Task.Run(() => RunHelper(shadow), CancellationToken.None);
                    }
                }

                internal int FirstRow { get; }
                internal int LastRow { get; }
#if NET8_0_OR_GREATER
                internal bool DecodeValues => true;
#else
                internal bool DecodeValues => false;
#endif
                internal bool StopRequested => _stopRequested;

                /// <summary>The first byte of the next chunk the sequential indexer has not reached yet.</summary>
                internal int NextBoundary => _nextChunk < _chunks.Length ? ChunkStart(_nextChunk) : int.MaxValue;

                private int ChunkStart(int chunk) => _firstChunkStart + chunk * _chunkBytes;

                private int ChunkEnd(int chunk) => chunk == _chunks.Length - 1
                    ? int.MaxValue : _firstChunkStart + (chunk + 1) * _chunkBytes;

                internal bool IsComplete(int available) => _inflation == null || available >= _expectedLength;

                internal int WaitForBytes(int previous, CancellationToken ct) =>
                    _inflation == null ? _bufferedLength : _inflation.WaitForMore(previous, ct);

                /// <summary>
                /// Returns the completed helper chunk whose boundary the sequential indexer just reached,
                /// or null after claiming an unstarted chunk for sequential indexing.
                /// </summary>
                internal ParallelRowChunk? TakeChunkAt(int position, CancellationToken ct) {
                    while (_nextChunk < _chunks.Length && position >= ChunkStart(_nextChunk)) {
                        ParallelRowChunk chunk = _chunks[_nextChunk++];
                        if (Interlocked.CompareExchange(ref chunk.State, ParallelRowChunk.OwnedBySequentialIndexer,
                                ParallelRowChunk.Unclaimed) == ParallelRowChunk.Unclaimed) return null;
                        if (Volatile.Read(ref chunk.State) != ParallelRowChunk.Completed) {
                            lock (_gate) {
                                while (Volatile.Read(ref chunk.State) != ParallelRowChunk.Completed) {
                                    ct.ThrowIfCancellationRequested();
                                    Monitor.Wait(_gate, 10);
                                }
                            }
                        }
                        // A later chunk boundary inside this chunk's rows is handled by the next call.
                        if (chunk.RowCount > 0 && chunk.Start == position) return chunk;
                        chunk.ReturnStorage();
                    }
                    return null;
                }

                private void RunHelper(ExcelUtf8RangeRowSource shadow) {
                    CancellationToken token = _stop.Token;
                    while (!_stopRequested) {
                        int claim = Interlocked.Increment(ref _nextClaim);
                        if (claim >= _chunks.Length) return;
                        ParallelRowChunk chunk = _chunks[claim];
                        if (Interlocked.CompareExchange(ref chunk.State, ParallelRowChunk.Running,
                                ParallelRowChunk.Unclaimed) != ParallelRowChunk.Unclaimed) continue;
                        try {
                            shadow.IndexParallelChunk(this, chunk, ChunkStart(claim), ChunkEnd(claim), token);
                        } catch (Exception) {
                            // Cancellation or an unexpected failure leaves the chunk to the sequential indexer.
                            chunk.RowCount = 0;
                        } finally {
                            lock (_gate) {
                                Volatile.Write(ref chunk.State, ParallelRowChunk.Completed);
                                Monitor.PulseAll(_gate);
                            }
                        }
                    }
                }

                public void Dispose() {
                    _stopRequested = true;
                    _stop.Cancel();
                    foreach (Task helper in _helpers) {
                        try { helper.GetAwaiter().GetResult(); } catch (Exception) { }
                    }
                    foreach (ParallelRowChunk chunk in _chunks) chunk.ReturnStorage();
                    _stop.Dispose();
                }
            }

            private ExcelUtf8RangeRowSource CreateParallelIndexShadow() =>
                new ExcelUtf8RangeRowSource(_owner, _buffer!, _length) {
                    _ownsBuffer = false,
                    _isParallelIndexHelper = true,
                    _firstColumn = _firstColumn,
                    _fieldCount = _fieldCount,
                    _repeatedRowSuffixStart = _repeatedRowSuffixStart,
                    _repeatedRowSuffixLength = _repeatedRowSuffixLength,
                    _repeatedRowIsEmpty = _repeatedRowIsEmpty
                };
        }
    }
}
