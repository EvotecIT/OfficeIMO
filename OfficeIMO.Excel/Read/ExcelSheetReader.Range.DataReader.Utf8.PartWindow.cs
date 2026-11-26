#nullable enable

using System.Buffers;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            // The window keeps the worksheet head (everything before the first row) at the start of the buffer, then
            // the bytes that hot-loop caches compare against, then a sliding run of rows. A row byte at buffer index i
            // (at or after the sliding run) is at part offset i + _windowShift. Row positions retained for delivery are
            // part offsets, so they stay valid when the window slides.
            private Stream? _partStream;
            private bool _windowedPart;
            private int _windowFilled;
            private int _windowShift;

            /// <summary>
            /// Qualifies a large worksheet with the same complete proof as the buffered readers, while holding only
            /// a bounded window of it. Rows are then delivered from a second inflation of the part, so the first row is
            /// returned only after the whole worksheet has been checked, and peak memory does not grow with the part.
            /// </summary>
            private static bool TryCreateWindowedUsedRange(ExcelSheetReader owner, CancellationToken ct,
                out ExcelUtf8RangeRowSource? source, out int declaredFirstColumn) {
                source = null;
                declaredFirstColumn = 0;
                OpenXmlPackagePartBufferReader? reader = owner._partBufferReader;
                if (!WindowedReadsSupported || reader == null || owner._opt.DisplayText != null
                    || !reader.TryGetLength(owner._worksheetPartName, out long length)
                    || length < owner._opt.WindowedWorksheetMinimumBytes || length > int.MaxValue
                    || !reader.TryOpenExactPartStream(owner._worksheetPartName, ct, out Stream? stream)) return false;

                byte[]? window = null;
                Stream partStream = stream!;
                ExcelUtf8RangeRowSource candidate;
                try {
                    window = OpenXmlPartBufferPool.Rent(owner._opt.WorksheetWindowBytes);
                    partStream = WrapPartStream(reader, stream!, ct);
                    candidate = new ExcelUtf8RangeRowSource(owner, window, 0) { _partStream = partStream, _windowedPart = true };
                } catch {
                    partStream.Dispose();
                    if (window != null) OpenXmlPartBufferPool.Return(window);
                    throw;
                }
                try {
                    while (candidate._windowFilled < 256 && candidate.TryReadIntoWindow(ct)) { }
                    if (!CanBufferUtf8WorksheetPrefix(candidate._buffer!, candidate._windowFilled)) {
                        candidate.Dispose();
                        return false;
                    }
                    int position = 0;
                    if (!candidate.TryAdvanceWindow(ref position, ct)
                        || !candidate.TryPrepareUsedRange(ct, allowUnknownRows: true, out declaredFirstColumn)) {
                        candidate.Dispose();
                        return false;
                    }
                    candidate.BeginWindowedDelivery(ct);
                    source = candidate;
                    return true;
                } catch {
                    candidate.Dispose();
                    throw;
                }
            }

            // Down-level text proof checks the whole buffer at the end rather than each block as it arrives, so it
            // needs the whole part.
#if NET8_0_OR_GREATER
            private static bool WindowedReadsSupported => true;
#else
            private static bool WindowedReadsSupported => false;
#endif

            /// <summary>
            /// Decompresses ahead on a worker only when the worksheet has its own archive cursor; otherwise another
            /// part read on this thread would share the cursor with the worker.
            /// </summary>
            private static Stream WrapPartStream(OpenXmlPackagePartBufferReader reader, Stream stream, CancellationToken ct) =>
                reader.AllowsPipelining ? ExcelPrefetchingReadStream.Wrap(stream, ct) : stream;

            /// <summary>Indexes every row into one reused metadata row; only row positions survive qualification.</summary>
            private void EnterWindowedCellMode() {
                if (!UsesCellWindow) PromoteProvisionalIndexToWindow();
                if (_rowContentStarts == null) {
                    _rowContentStarts = ExcelArrayPool.RentInts(_rowIndexes!.Length);
                    _rowContentEnds = ExcelArrayPool.RentInts(_rowIndexes.Length);
                }
            }

            private int GetWindowIndex(int partOffset) => _windowedPart ? partOffset - _windowShift : partOffset;

            private int GetPartOffset(int windowIndex) => _windowedPart ? windowIndex + _windowShift : windowIndex;

            private bool TryReadIntoWindow(CancellationToken ct) {
                if (_windowFilled == _buffer!.Length) GrowWindow();
                int read = _partStream!.Read(_buffer!, _windowFilled, _buffer!.Length - _windowFilled);
                // Record the bytes before observing cancellation, so a canceled refill can be retried.
                _windowFilled += read;
                ct.ThrowIfCancellationRequested();
                return read != 0;
            }

            private void GrowWindow() {
                byte[] previous = _buffer!;
                int length = checked(previous.Length * 2);
                // A row or trailing worksheet content larger than the pool's largest buffer still has to fit.
                byte[] next = length <= MaximumBufferSize ? OpenXmlPartBufferPool.Rent(length) : new byte[length];
                Buffer.BlockCopy(previous, 0, next, 0, _windowFilled);
                _buffer = next;
                OpenXmlPartBufferPool.Return(previous);
            }

            /// <summary>
            /// Makes more complete rows visible to the qualification indexer. Like the pipelined reader, it exposes
            /// bytes only through the last complete row and everything once the part has ended. Rows before
            /// <paramref name="position"/> have been indexed, so their bytes can leave the window.
            /// </summary>
            private bool TryAdvanceWindow(ref int position, CancellationToken ct) {
                if (_sheetDataContentStart >= 0 && position >= _length
                    && position - _sheetDataContentStart >= _buffer!.Length / 2) {
                    position -= CompactWindow(position);
                }
                while (true) {
                    if (_windowFilled == _buffer!.Length && _sheetDataContentStart >= 0
                        && position >= _length && position > WindowRowRegionStart()) {
                        position -= CompactWindow(position);
                    }
                    if (!TryReadIntoWindow(ct)) {
                        if (_windowFilled <= _length) return false;
                        _length = _windowFilled;
                        ValidateProducedCanonicalText(ct);
                        return true;
                    }
                    int rowEnd = FindLastRowEnd(_length, _windowFilled);
                    if (rowEnd >= 0) {
                        _length = rowEnd;
                        ValidateProducedCanonicalText(ct);
                        return true;
                    }
                    // Without a row end there is nothing to slide past, so stop before holding the whole part.
                    if (UnslidBytes() >= MaximumWindowHeadBytes) return false;
                }
            }

            private const int MaximumWindowHeadBytes = 64 * 1024 * 1024;

            /// <summary>
            /// Returns the end of the last complete row closing tag (<c>&lt;/row&gt;</c> or a prefixed
            /// <c>&lt;/x:row&gt;</c>) in the range, or -1.
            /// </summary>
            private int FindLastRowEnd(int start, int end) {
                byte[] buffer = _buffer!;
                int limit = end;
                while (limit > start) {
                    int index = buffer.AsSpan(start, limit - start).LastIndexOf("row>"u8);
                    if (index < 0) return -1;
                    int name = start + index;
                    int cursor = name - 1;
                    if (cursor >= start && buffer[cursor] == (byte)':') {
                        cursor--;
                        while (cursor >= start && buffer[cursor] != (byte)'/' && buffer[cursor] != (byte)'<'
                            && buffer[cursor] > (byte)' ' && buffer[cursor] != (byte)'>') cursor--;
                    }
                    if (cursor - 1 >= start && buffer[cursor] == (byte)'/' && buffer[cursor - 1] == (byte)'<') return name + 4;
                    limit = name;
                }
                return -1;
            }

            /// <summary>
            /// Reads the bytes after the last row, which the final envelope proof inspects. Declines, before any row
            /// is returned, when they would not fit the same bound as a run without a row end.
            /// </summary>
            private bool CompleteWindowedQualification(CancellationToken ct) {
                while (TryReadIntoWindow(ct)) {
                    if (UnslidBytes() >= MaximumWindowHeadBytes) return false;
                }
                if (_windowFilled > _length) {
                    _length = _windowFilled;
                    ValidateProducedCanonicalText(ct);
                }
                return true;
            }

            // Bytes held that the window cannot slide past: the head before the rows are found, then the bytes
            // after the pinned area.
            private int UnslidBytes() => _windowFilled - Math.Max(0, WindowRowRegionStart());

            // Bytes that hot-loop caches compare against are copied next to the head when rows leave the window.
            private int _windowPinnedLength;

            private int WindowRowRegionStart() => _sheetDataContentStart + _windowPinnedLength;

            /// <summary>
            /// Drops row bytes before <paramref name="keepFrom"/>, keeping the head and the bytes cached tag shapes
            /// refer to. Returns how far later bytes moved toward the start of the buffer.
            /// </summary>
            private int CompactWindow(int keepFrom) {
                int head = _sheetDataContentStart;
                byte[] buffer = _buffer!;
                int pinnedLength = 0;
                if (_repeatedRowSuffixStart >= head && _repeatedRowSuffixStart < keepFrom) pinnedLength += _repeatedRowSuffixLength;
                Utf8ImplicitCellTagCacheEntry[]? tags = _implicitCellTagCache;
                if (tags != null) {
                    for (int ordinal = 0; ordinal < _fieldCount; ordinal++) {
                        if (tags[ordinal].Length > 0 && tags[ordinal].Start >= head && tags[ordinal].Start < keepFrom) {
                            pinnedLength += tags[ordinal].Length;
                        }
                    }
                }
                int delta = keepFrom - head - pinnedLength;
                if (delta <= 0) return 0;
                int movable = Math.Max(0, _windowFilled - keepFrom);
                int skipped = Math.Max(0, keepFrom - _windowFilled);

                byte[] pinned = ArrayPool<byte>.Shared.Rent(Math.Max(1, pinnedLength));
                try {
                    int cursor = 0;
                    if (_repeatedRowSuffixStart >= head && _repeatedRowSuffixStart < keepFrom) {
                        Buffer.BlockCopy(buffer, _repeatedRowSuffixStart, pinned, cursor, _repeatedRowSuffixLength);
                        _repeatedRowSuffixStart = head + cursor;
                        cursor += _repeatedRowSuffixLength;
                    } else if (_repeatedRowSuffixStart >= keepFrom) {
                        _repeatedRowSuffixStart -= delta;
                    }
                    if (tags != null) {
                        for (int ordinal = 0; ordinal < _fieldCount; ordinal++) {
                            Utf8ImplicitCellTagCacheEntry tag = tags[ordinal];
                            if (tag.Length <= 0) continue;
                            int start;
                            if (tag.Start >= head && tag.Start < keepFrom) {
                                Buffer.BlockCopy(buffer, tag.Start, pinned, cursor, tag.Length);
                                start = head + cursor;
                                cursor += tag.Length;
                            } else if (tag.Start >= keepFrom) {
                                start = tag.Start - delta;
                            } else {
                                continue;
                            }
                            tags[ordinal] = new Utf8ImplicitCellTagCacheEntry(start, tag.Length, tag.Kind, tag.StyleIndex,
                                tag.IsEmpty, tag.EncodedKind, tag.DateStylesEnabled);
                        }
                    }
                    if (movable > 0) Buffer.BlockCopy(buffer, keepFrom, buffer, head + pinnedLength, movable);
                    Buffer.BlockCopy(pinned, 0, buffer, head, pinnedLength);
                } finally {
                    ArrayPool<byte>.Shared.Return(pinned, clearArray: true);
                }
                int filled = head + pinnedLength + movable;
                // Dropped bytes may hold customer data; clear the vacated tail as pooled buffers are cleared on return.
                Array.Clear(buffer, filled, _windowFilled - filled);
                _windowPinnedLength = pinnedLength;
                _windowShift += delta;
                _windowFilled = filled;
                _length = Math.Max(head + pinnedLength, _length - delta);
                // The envelope check copies the markup after sheetData from this offset.
                if (_sheetDataEndTagStart >= keepFrom) _sheetDataEndTagStart -= delta;
#if NET8_0_OR_GREATER
                _validatedCanonicalTextLength = Math.Max(head + pinnedLength, _validatedCanonicalTextLength - delta);
#endif
                // Delivery can move past rows it skips without reading them.
                SkipPartBytes(skipped);
                // Cached strings hold buffer offsets of their source text.
                if (_stringCache != null) Array.Clear(_stringCache, 0, _stringCache.Length);
                return delta;
            }

            /// <summary>
            /// Restarts the part for delivery. The head must match the qualified head byte for byte; rows are later
            /// re-indexed from their qualified part offsets and must end exactly where qualification saw them end.
            /// </summary>
            private void BeginWindowedDelivery(CancellationToken ct) {
                _partStream!.Dispose();
                _partStream = null;
                // A self-closing sheetData has no rows to deliver.
                if (_sheetDataContentStart < 0) return;
                // Delivery observes cancellation between windows (see SelectCellWindowRow), not inside a refill.
                if (!_owner._partBufferReader!.TryOpenExactPartStream(_owner._worksheetPartName, CancellationToken.None, out Stream? stream)) {
                    throw new InvalidDataException($"Worksheet '{_owner._sheetName}' could not be reopened for reading.");
                }
                try {
                    _partStream = WrapPartStream(_owner._partBufferReader!, stream!, CancellationToken.None);
                } catch {
                    stream!.Dispose();
                    throw;
                }
                int head = _sheetDataContentStart;
                byte[] check = ArrayPool<byte>.Shared.Rent(Math.Min(Math.Max(head, 1), 64 * 1024));
                try {
                    int compared = 0;
                    while (compared < head) {
                        int read = _partStream!.Read(check, 0, Math.Min(check.Length, head - compared));
                        if (read == 0 || !check.AsSpan(0, read).SequenceEqual(_buffer.AsSpan(compared, read))) {
                            throw new InvalidDataException($"Worksheet '{_owner._sheetName}' changed while it was being read.");
                        }
                        compared += read;
                    }
                } finally {
                    ArrayPool<byte>.Shared.Return(check, clearArray: true);
                }
                Array.Clear(_buffer!, head, _windowFilled - head);
                _windowFilled = head;
                _length = head;
                _windowShift = 0;
                _windowPinnedLength = 0;
                _repeatedRowSuffixStart = -1;
                if (_implicitCellTagCache != null) Array.Clear(_implicitCellTagCache, 0, _implicitCellTagCache.Length);
                _cachedImplicitRowReady = false;
                _cachedImplicitRowRejected = false;
                if (_stringCache != null) Array.Clear(_stringCache, 0, _stringCache.Length);
            }

            /// <summary>
            /// Loads the bytes of up to <paramref name="count"/> qualified rows starting at
            /// <paramref name="physicalRow"/> and returns how many of them fit the window.
            /// </summary>
            private int LoadWindowedDeliveryRows(int physicalRow, int count, CancellationToken ct) {
                int keep = _rowContentStarts![physicalRow];
                if (_selectedPhysicalRow >= 0 && _selectedPhysicalRow < physicalRow) {
                    // A canceled refill restores the selected row from the window.
                    keep = Math.Min(keep, _rowContentStarts[_selectedPhysicalRow]);
                }
                int keepIndex = keep - _windowShift;
                if (keepIndex < WindowRowRegionStart()) {
                    throw new InvalidOperationException("A qualified worksheet row is no longer in the read window.");
                }
                if (keepIndex - _sheetDataContentStart >= _buffer!.Length / 2
                    || _rowContentEnds![physicalRow + count - 1] - _windowShift > _buffer.Length) {
                    CompactWindow(keepIndex);
                }
                // Rows are in part order, so the rows that fit are a prefix of the requested run.
                int capacityEnd = _buffer!.Length + _windowShift;
                int low = 1, high = count;
                while (low < high) {
                    int middle = low + (high - low + 1) / 2;
                    if (_rowContentEnds![physicalRow + middle - 1] <= capacityEnd) low = middle;
                    else high = middle - 1;
                }
                count = low;
                int required = _rowContentEnds![physicalRow + count - 1] - _windowShift;
                while (_windowFilled < required) {
                    if (!TryReadIntoWindow(ct)) {
                        throw new InvalidDataException($"Worksheet '{_owner._sheetName}' changed while it was being read.");
                    }
                }
                _length = _windowFilled;
                return count;
            }

            private void SkipPartBytes(int count) {
                byte[] buffer = _buffer!;
                while (count > 0) {
                    int free = buffer.Length - _windowFilled;
                    int read = _partStream!.Read(buffer, _windowFilled, Math.Min(free, count));
                    if (read == 0) throw new InvalidDataException($"Worksheet '{_owner._sheetName}' changed while it was being read.");
                    Array.Clear(buffer, _windowFilled, read);
                    count -= read;
                }
            }

            private void DisposePartWindow() {
                _partStream?.Dispose();
                _partStream = null;
            }
        }
    }
}
