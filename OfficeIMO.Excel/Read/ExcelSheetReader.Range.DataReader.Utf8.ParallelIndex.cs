#nullable enable

using System.Buffers;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const int MinimumParallelIndexBytes = 1024 * 1024;
            private const int MaximumParallelIndexHelpers = 3;
            private ParallelRowIndex? _parallelIndex;
            // Helpers share package cursors with the sequential indexer, so they must not load parts.
            private bool _isParallelIndexHelper;
            private bool _parallelIndexAttempted;
            // The complete part length when indexing overlaps inflation; zero uses the buffered length.
            private int _expectedLength;

            /// <summary>
            /// Starts helper workers that index later byte chunks of a uniform worksheet into
            /// private metadata. The sequential indexer remains authoritative: it adopts helper rows
            /// only when a chunk's first row starts exactly where its own parse arrived, so a helper
            /// can only repeat work the sequential pass would do with the same row shape. It never
            /// waits for a helper; unfinished chunk tails are indexed sequentially.
            /// </summary>
            private void TryStartParallelIndex(int firstRow, int lastRow, int position) {
                if (_parallelIndexAttempted) return;
                _parallelIndexAttempted = true;
                int expected = _expectedLength > 0 ? _expectedLength : _length;
#if !NET8_0_OR_GREATER
                // On .NET Framework and .NET Standard hosts the helpers measured slower than the
                // sequential indexer, so those builds keep it.
                _ = expected;
#else
                if (expected < MinimumParallelIndexBytes || UsesCellWindow || _formulaStarts != null
                    || !_options.Execution.AllowsBackgroundProcessing) return;
                int helpers = Math.Min(MaximumParallelIndexHelpers, Environment.ProcessorCount - 2);
                int? maximumDegree = _options.Execution.MaxDegreeOfParallelism;
                if (maximumDegree.HasValue) helpers = Math.Min(helpers, maximumDegree.Value - 1);
                if (helpers <= 0) return;
                int averageRowBytes = Math.Max(32, (position - Math.Max(0, _sheetDataContentStart)) / Math.Max(1, _rowCount));
                // A sheet expected to outgrow the whole-sheet cell budget is read through a rolling
                // window; its index metadata is not retained, so helpers would only duplicate work.
                // Adoption enforces the exact budget, so this estimate only screens clear cases.
                if (_unknownRowExtent
                    && (long)expected / averageRowBytes * _fieldCount > Math.Min(MaximumIndexedCells, _options.MaxDataReaderBufferedCells) * 5L / 4) return;
                helpers = ExcelParallelReadBudget.Reserve(helpers);
                if (helpers <= 0) return;
                int chunkBytes = Math.Max(256 * 1024, Math.Min(4 * 1024 * 1024, expected / (helpers * 8)));
                try {
                    _parallelIndex = new ParallelRowIndex(this, expected, chunkBytes, firstChunkStart: position,
                        averageRowBytes, firstRow, lastRow, helpers, _options.CancellationToken);
                } catch {
                    ExcelParallelReadBudget.Release(helpers);
                    throw;
                }
#endif
            }
            /// <summary>Adopts the completed helper rows that start at the current row boundary.</summary>
            private bool TryStitchParallelChunk(ref int position, ref int previousRow, CancellationToken ct) {
                ParallelRowIndex index = _parallelIndex!;
                ParallelRowChunk? chunk = index.TakeChunkAt(position);
                if (chunk == null) return false;
                bool completed = Volatile.Read(ref chunk.State) == ParallelRowChunk.Completed
                    || WaitWhileHelperProgresses(chunk);
                // Rows published before this read are immutable; a running helper stops at its next row.
                int rows = Volatile.Read(ref chunk.CompletedRows);
                if (!completed) Volatile.Write(ref chunk.Abandoned, true);
                try {
                    // Attribute-free rows continue the sequential numbering at this boundary.
                    int rowBase = chunk.ImplicitRows ? previousRow : 0;
                    if (rows == 0 || chunk.Start != position || UsesCellWindow
                        || chunk.RowIndexes![0] + rowBase <= previousRow
                        || chunk.ImplicitRows && (rowBase + 1 < index.FirstRow || (long)rowBase + rows > index.LastRow)) return false;
                    // A provisional index that would outgrow its cell budget becomes a rolling window;
                    // the sequential indexer performs that promotion.
                    if (_unknownRowExtent
                        && (long)(_rowCount + rows) * _fieldCount > Math.Min(MaximumIndexedCells, _options.MaxDataReaderBufferedCells)) return false;
                    int end = chunk.RowContentEnds![rows - 1];
                    while (_length < end) {
                        if (!TryAdvanceInflation(ct)) return false;
                    }

                    int cells = checked(rows * _fieldCount);
                    EnsureRowCapacity(_rowCount + rows);
                    int rowOffset = GetQualificationRowOffset(_rowCount);
                    if (rowBase == 0) {
                        Array.Copy(chunk.RowIndexes, 0, _rowIndexes!, _rowCount, rows);
                    } else {
                        for (int row = 0; row < rows; row++) _rowIndexes![_rowCount + row] = chunk.RowIndexes[row] + rowBase;
                    }
                    if (_unknownRowExtent) {
                        Array.Copy(chunk.RowContentStarts!, 0, _rowContentStarts!, _rowCount, rows);
                        Array.Copy(chunk.RowContentEnds, 0, _rowContentEnds!, _rowCount, rows);
                    }
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

                    for (int row = 0; row < rows; row++) {
                        int lastColumn = chunk.RowLastColumns![row];
                        if (lastColumn <= 0) continue;
                        int rowIndex = chunk.RowIndexes[row] + rowBase;
                        if (_minimumCellRow == int.MaxValue) _minimumCellRow = rowIndex;
                        _maximumCellRow = rowIndex;
                        if (lastColumn > _maximumCellColumn) _maximumCellColumn = lastColumn;
                        int firstColumn = chunk.RowFirstColumns![row];
                        if (firstColumn < _minimumCellColumn) _minimumCellColumn = firstColumn;
                    }

                    if (chunk.RequiresFullValidation) _sheetDataSupportsFastValidation = false;
                    previousRow = chunk.RowIndexes[rows - 1] + rowBase;
                    _rowCount += rows;
                    position = end;
                    return true;
                } finally {
                    // A running helper may still write later rows; its storage returns when helpers are joined.
                    if (completed) chunk.ReturnStorage();
                }
            }

            /// <summary>
            /// Waits for a running chunk only while its helper keeps publishing rows. A helper that
            /// stalls, for example on a busy host, is abandoned after a short pause and the
            /// sequential indexer adopts its published prefix instead.
            /// </summary>
            private static bool WaitWhileHelperProgresses(ParallelRowChunk chunk) {
                long stallTicks = System.Diagnostics.Stopwatch.Frequency / 5000;
                int published = Volatile.Read(ref chunk.CompletedRows);
                long lastProgress = System.Diagnostics.Stopwatch.GetTimestamp();
                var spinner = new SpinWait();
                while (true) {
                    if (Volatile.Read(ref chunk.State) == ParallelRowChunk.Completed) return true;
                    int current = Volatile.Read(ref chunk.CompletedRows);
                    long now = System.Diagnostics.Stopwatch.GetTimestamp();
                    if (current != published) {
                        published = current;
                        lastProgress = now;
                    } else if (now - lastProgress > stallTicks) {
                        return false;
                    }
#if NET8_0_OR_GREATER
                    spinner.SpinOnce(sleep1Threshold: -1);
#else
                    if (spinner.NextSpinWillYield) Thread.Yield(); else spinner.SpinOnce();
#endif
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
                    if (Volatile.Read(ref chunk.Abandoned)) return;
                    int searchEnd = Math.Min(available, chunkEnd);
                    byte[] rowToken = RowStartToken();
                    while (searchFrom < searchEnd) {
                        int relative = _buffer!.AsSpan(searchFrom, searchEnd - searchFrom).IndexOf(rowToken);
                        if (relative < 0) {
                            searchFrom = Math.Max(searchFrom, searchEnd - rowToken.Length + 1);
                            break;
                        }
                        int candidate = searchFrom + relative;
                        if (candidate + rowToken.Length >= available) {
                            searchFrom = candidate;
                            break;
                        }
                        if (_buffer![candidate + rowToken.Length] is (byte)' ' or (byte)'>') {
                            position = candidate;
                            break;
                        }
                        searchFrom = candidate + rowToken.Length;
                    }
                    if (position >= 0) break;
                    if (available >= chunkEnd || index.IsComplete(available)) return;
                    int next = index.WaitForBytes(available, ct);
                    if (next <= available) return;
                    available = next;
                }

                _length = available;
                _sheetDataSupportsFastValidation = true;
                chunk.Start = position;
                // Attribute-free rows take their numbers from the sequential indexer when adopted;
                // helpers number them from one. Other rows must repeat the validated row shape.
                bool implicitRows = IsAttributeFreeRowTag(position);
                chunk.ImplicitRows = implicitRows;
                // Storage is fixed before the first published row, so the sequential indexer can copy
                // published rows while this helper continues.
                int capacity = chunk.Allocate(_fieldCount, index.DecodeValues, index.RecordRowPositions);
                _rowIndexes = chunk.RowIndexes;
                _valueStarts = chunk.ValueStarts;
                _valueLengths = chunk.ValueLengths;
                _cellKinds = chunk.CellKinds;
                int rows = 0;
                int previousRow = 0;
                while (position < chunkEnd && rows < capacity && !index.StopRequested
                    && !Volatile.Read(ref chunk.Abandoned)) {
                    if (position >= _length) {
                        int next = index.WaitForBytes(_length, ct);
                        if (next <= _length) break;
                        _length = next;
                    }

                    SkipWhitespaceBetweenRows(ref position);
                    int rowStart = position;
                    int rowIndex;
                    bool emptyRow;
                    if (implicitRows) {
                        if (!IsAttributeFreeRowTag(position)) break;
                        position += 5 + _worksheetPrefixLength;
                        rowIndex = rows + 1;
                        emptyRow = false;
                    } else if (!TryReadRepeatedRowStartTag(ref position, out Utf8Tag tag, out rowIndex)
                        || rowIndex <= previousRow || rowIndex < index.FirstRow || rowIndex > index.LastRow) {
                        break;
                    } else {
                        emptyRow = tag.IsEmpty;
                    }

                    int contentStart = position;
                    _rowCount = rows;
                    _minimumCellColumn = int.MaxValue;
                    _maximumCellColumn = 0;
                    _maximumCellRow = 0;
                    int rowOffset = rows * _fieldCount;
                    if (!implicitRows) InitializeMetadataRow(rowOffset);
                    if (!emptyRow) {
                        bool indexed;
                        try {
                            // The implicit indexer initializes every slot of a committed row itself.
                            indexed = implicitRows
                                ? TryIndexCanonicalImplicitDenseRow(ref position, rowIndex, rowOffset, ct)
                                : TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
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
                    if (chunk.RowContentStarts != null) chunk.RowContentStarts[rows] = contentStart;
                    chunk.RowContentEnds![rows] = position;
                    chunk.RowFirstColumns![rows] = _minimumCellColumn;
                    chunk.RowLastColumns![rows] = _maximumCellRow > 0 ? _maximumCellColumn : 0;
                    chunk.RowIndexes![rows] = rowIndex;
                    rows++;
                    // Published with the row so an adopted prefix carries its validation requirement.
                    if (!_sheetDataSupportsFastValidation) chunk.RequiresFullValidation = true;
                    Volatile.Write(ref chunk.CompletedRows, rows);
                    previousRow = rowIndex;
                }
            }

            private sealed class ParallelIndexRowDeferredException : Exception { }

            private sealed class ParallelRowChunk {
                internal const int Unclaimed = 0;
                internal const int Running = 1;
                internal const int Completed = 2;
                internal const int OwnedBySequentialIndexer = 3;

                internal int State;
                internal int CompletedRows;
                internal bool Abandoned;
                internal int Start = -1;
                internal int EstimatedRows;
                internal bool ImplicitRows;
                internal bool RequiresFullValidation;
                internal int[]? RowIndexes;
                internal int[]? RowContentStarts;
                internal int[]? RowContentEnds;
                internal int[]? RowFirstColumns;
                internal int[]? RowLastColumns;
                internal int[]? ValueStarts;
                internal int[]? ValueLengths;
                internal byte[]? CellKinds;
                internal long[]? DecodedValues;
                internal byte[]? DecodedTags;

                /// <summary>Rents fixed storage for the chunk and returns its row capacity.</summary>
                internal int Allocate(int fieldCount, bool decodeValues, bool recordStarts) {
                    int rows = EstimatedRows;
                    int cells = checked(rows * fieldCount);
                    RowIndexes = ArrayPool<int>.Shared.Rent(rows);
                    if (recordStarts) RowContentStarts = ArrayPool<int>.Shared.Rent(rows);
                    RowContentEnds = ArrayPool<int>.Shared.Rent(rows);
                    RowFirstColumns = ArrayPool<int>.Shared.Rent(rows);
                    RowLastColumns = ArrayPool<int>.Shared.Rent(rows);
                    ValueStarts = ArrayPool<int>.Shared.Rent(cells);
                    ValueLengths = ArrayPool<int>.Shared.Rent(cells);
                    CellKinds = ArrayPool<byte>.Shared.Rent(cells);
                    if (decodeValues) {
                        DecodedValues = ArrayPool<long>.Shared.Rent(cells);
                        DecodedTags = ArrayPool<byte>.Shared.Rent(cells);
                    }
                    return rows;
                }

                internal void ReturnStorage() {
                    Return(ref RowIndexes);
                    Return(ref RowContentStarts);
                    Return(ref RowContentEnds);
                    Return(ref RowFirstColumns);
                    Return(ref RowLastColumns);
                    Return(ref ValueStarts);
                    Return(ref ValueLengths);
                    if (CellKinds != null) ArrayPool<byte>.Shared.Return(CellKinds);
                    if (DecodedValues != null) ArrayPool<long>.Shared.Return(DecodedValues);
                    if (DecodedTags != null) ArrayPool<byte>.Shared.Return(DecodedTags);
                    CellKinds = DecodedTags = null;
                    DecodedValues = null;
                }

                private static void Return(ref int[]? values) {
                    if (values != null) ArrayPool<int>.Shared.Return(values);
                    values = null;
                }
            }

            private sealed class ParallelRowIndex : IDisposable {
                private readonly WorksheetInflationProgress? _inflation;
                private readonly int _bufferedLength;
                private readonly int _expectedLength;
                private readonly int _chunkBytes;
                private readonly int _firstChunkStart;
                private readonly int _reservedHelpers;
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
                    RecordRowPositions = owner._unknownRowExtent;
                    _bufferedLength = owner._length;
                    _expectedLength = expectedLength;
                    _chunkBytes = chunkBytes;
                    _reservedHelpers = helpers;
                    // Chunk zero begins at the current row and always belongs to the sequential indexer.
                    _firstChunkStart = firstChunkStart;
                    FirstRow = firstRow;
                    LastRow = lastRow;
                    int remaining = Math.Max(0, expectedLength - firstChunkStart);
                    int count = Math.Max(1, (remaining + chunkBytes - 1) / chunkBytes);
                    _chunks = new ParallelRowChunk[count];
                    // A quarter above the expected rows; a chunk of unusually short rows stops early
                    // and the sequential indexer continues from the last published row.
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
                // Provisional whole-sheet indexes keep each row's byte range for a later window.
                internal bool RecordRowPositions { get; }
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
                /// Returns the started helper chunk whose boundary the sequential indexer just reached,
                /// or null after claiming an unstarted chunk for sequential indexing. Never waits.
                /// </summary>
                internal ParallelRowChunk? TakeChunkAt(int position) {
                    ParallelRowChunk? found = null;
                    while (_nextChunk < _chunks.Length && position >= ChunkStart(_nextChunk)) {
                        ParallelRowChunk chunk = _chunks[_nextChunk++];
                        if (Interlocked.CompareExchange(ref chunk.State, ParallelRowChunk.OwnedBySequentialIndexer,
                                ParallelRowChunk.Unclaimed) == ParallelRowChunk.Unclaimed) {
                            if (found != null) Volatile.Write(ref found.Abandoned, true);
                            found = null;
                            continue;
                        }
                        // A boundary the sequential indexer already passed cannot be adopted.
                        if (found != null) Volatile.Write(ref found.Abandoned, true);
                        found = chunk;
                    }
                    return found;
                }

                private void RunHelper(ExcelUtf8RangeRowSource shadow) {
                    try {
                        RunHelperChunks(shadow);
                    } finally {
                        shadow.ReturnImplicitCellTagCache();
                    }
                }

                private void RunHelperChunks(ExcelUtf8RangeRowSource shadow) {
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
                            // Cancellation or an unexpected failure ends the chunk at its last published row.
                        } finally {
                            Volatile.Write(ref chunk.State, ParallelRowChunk.Completed);
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
                    ExcelParallelReadBudget.Release(_reservedHelpers);
                }
            }

            private byte[]? _rowStartToken;

            /// <summary>The bytes that open a row element in this worksheet's prefix, without the name terminator.</summary>
            private byte[] RowStartToken() {
                if (_rowStartToken != null) return _rowStartToken;
                var token = new byte[4 + _worksheetPrefixLength];
                token[0] = (byte)'<';
                Array.Copy(_buffer!, _worksheetPrefixStart, token, 1, _worksheetPrefixLength);
                token[token.Length - 3] = (byte)'r';
                token[token.Length - 2] = (byte)'o';
                token[token.Length - 1] = (byte)'w';
                return _rowStartToken = token;
            }

            private ExcelUtf8RangeRowSource CreateParallelIndexShadow() =>
                new ExcelUtf8RangeRowSource(_owner, _buffer!, _length) {
                    _ownsBuffer = false,
                    _isParallelIndexHelper = true,
                    _worksheetPrefixStart = _worksheetPrefixStart,
                    _worksheetPrefixLength = _worksheetPrefixLength,
                    _firstColumn = _firstColumn,
                    _fieldCount = _fieldCount,
                    _repeatedRowSuffixStart = _repeatedRowSuffixStart,
                    _repeatedRowSuffixLength = _repeatedRowSuffixLength,
                    _repeatedRowIsEmpty = _repeatedRowIsEmpty
                };
        }
    }
}
