#nullable enable

using System.Buffers;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const int MinimumRowBatchPartBytes = 1024 * 1024;
            private const int RowBatchRows = 4096;

            // Producer side: the qualification indexer copies each finished row into the filling batch.
            private RowBatchChannel? _batchSink;
            // Reader side: rows come from batches handed over by a background qualification pass.
            private RowBatchChannel? _batchSource;
            private ExcelUtf8RangeRowSource? _batchProducer;
            private Task? _batchProducerTask;
            private CancellationTokenSource? _batchStop;
            private RowBatch? _currentBatch;
            private int _batchCursor;
            private RowBatchRange _batchRange;

            internal bool RowBatchesEnded { get; private set; }

            private bool UsesRowBatches => _batchSource != null;

            /// <summary>
            /// Validates the worksheet while rows are read: a background pass inflates and qualifies the part once,
            /// and the reader receives each qualified run of rows with its own copy of their bytes. A malformed
            /// row is reported by the read that reaches it.
            /// </summary>
            private static bool TryCreateRowBatchUsedRange(ExcelSheetReader owner, CancellationToken ct,
                out ExcelUtf8RangeRowSource? source, out int declaredFirstColumn) {
                source = null;
                declaredFirstColumn = 0;
                OpenXmlPackagePartBufferReader? reader = owner._partBufferReader;
                if (!WindowedReadsSupported || owner._opt.ValidateWorksheetBeforeFirstRow || owner._opt.DisplayText != null
                    || owner._opt.MaxDataReaderChunkRows <= 0
                    || reader?.AllowsPipelining != true
                    || !reader.TryGetLength(owner._worksheetPartName, out long length)
                    || length < MinimumRowBatchPartBytes || length > int.MaxValue
                    || HeadHasNoDimension(reader, owner._worksheetPartName, ct)) return false;
                // The background pass validates styles; load them here so only this thread builds them.
                _ = owner.Styles;
                if (!reader.TryOpenExactPartStream(owner._worksheetPartName, ct, out Stream? stream)) return false;
                byte[]? window = null;
                ExcelUtf8RangeRowSource producer;
                try {
                    window = OpenXmlPartBufferPool.Rent(owner._opt.WorksheetWindowBytes);
                    producer = new ExcelUtf8RangeRowSource(owner, window, 0) {
                        _partStream = WrapPartStream(reader, stream!, ct), _windowedPart = true
                    };
                } catch {
                    stream!.Dispose();
                    if (window != null) OpenXmlPartBufferPool.Return(window);
                    throw;
                }
                var channel = new RowBatchChannel();
                producer._batchSink = channel;
                var stop = CancellationTokenSource.CreateLinkedTokenSource(ct);
                Task task = Task.Run(() => producer.RunRowBatchProducer(channel, stop.Token), CancellationToken.None);
                bool started;
                RowBatchRange range;
                try {
                    started = channel.WaitForFirstRows(ct, out range);
                } catch {
                    StopProducer(channel, stop, task, producer);
                    throw;
                }
                if (!started) {
                    // Nothing was returned yet, so the ordinary readers still decide this worksheet.
                    StopProducer(channel, stop, task, producer);
                    return false;
                }

                var consumer = new ExcelUtf8RangeRowSource(owner, Array.Empty<byte>(), 0) {
                    _ownsBuffer = false,
                    _firstColumn = range.FirstColumn,
                    _fieldCount = range.FieldCount,
                    _batchSource = channel,
                    _batchProducer = producer,
                    _batchProducerTask = task,
                    _batchStop = stop,
                    _batchRange = range
                };
                declaredFirstColumn = range.FirstColumn;
                source = consumer;
                return true;
            }

            private const int DimensionProbeBytes = 16 * 1024;

            /// <summary>
            /// Reads the start of the worksheet: rows that begin without a declared dimension are validated before the
            /// first row, so no background pass is started for them. A head too long to probe is left to that pass.
            /// </summary>
            private static bool HeadHasNoDimension(OpenXmlPackagePartBufferReader reader, string partName, CancellationToken ct) {
                if (!reader.TryOpenExactPartStream(partName, ct, out Stream? stream)) return false;
                byte[] head = ArrayPool<byte>.Shared.Rent(DimensionProbeBytes);
                try {
                    int read = 0;
                    using (stream) {
                        int count;
                        while (read < DimensionProbeBytes && (count = stream!.Read(head, read, DimensionProbeBytes - read)) > 0) read += count;
                    }
                    int rows = IndexOfElement(head, read, "sheetData"u8);
                    return rows >= 0 && IndexOfElement(head, rows, "dimension"u8) < 0;
                } finally {
                    ArrayPool<byte>.Shared.Return(head);
                }
            }

            // The first start tag with this local name, with or without a prefix, before `end`.
            private static int IndexOfElement(byte[] bytes, int end, ReadOnlySpan<byte> localName) {
                for (int position = 1; position + localName.Length < end; position++) {
                    byte before = bytes[position - 1];
                    if ((before == (byte)'<' || before == (byte)':') && bytes.AsSpan(position, localName.Length).SequenceEqual(localName)) {
                        byte after = bytes[position + localName.Length];
                        if (after == (byte)'>' || after == (byte)'/' || after == (byte)' ' || after == (byte)'\t'
                            || after == (byte)'\r' || after == (byte)'\n') return position;
                    }
                }
                return -1;
            }

            private static void StopProducer(RowBatchChannel channel, CancellationTokenSource stop, Task task, ExcelUtf8RangeRowSource producer) {
                stop.Cancel();
                channel.Cancel();
                try { task.GetAwaiter().GetResult(); } catch { }
                stop.Dispose();
                producer.Dispose();
                channel.Release();
            }

            private void RunRowBatchProducer(RowBatchChannel channel, CancellationToken ct) {
                _batchCt = ct;
                try {
                    while (_windowFilled < 256 && TryReadIntoWindow(ct)) { }
                    int position = 0;
                    if (!CanBufferUtf8WorksheetPrefix(_buffer!, _windowFilled) || !TryAdvanceWindow(ref position, ct)) {
                        channel.Decline();
                        return;
                    }
                    bool qualified = TryPrepareUsedRange(ct, allowUnknownRows: true, out _);
                    channel.Flush();
                    // Content the compact indexer does not handle is read on by the general XML reader.
                    if (!qualified) throw new RowBatchHandoverException();
                    ThrowIfCellsOutsideDeclaredRange();
                    channel.Complete();
                } catch (RowBatchDeclinedException) {
                    channel.Decline();
                } catch (XlsxTabularFastPathNotSupportedException) {
                    channel.Flush();
                    channel.Fault(new RowBatchHandoverException());
                } catch (Exception exception) {
                    channel.Flush();
                    channel.Fault(exception);
                } finally {
                    _batchXmlCheck?.Dispose();
                }
            }

            /// <summary>Called by the qualification pass after the reader's columns are fixed.</summary>
            private void PublishRowBatchRange(int firstRow, int lastRow, int firstColumn, int fieldCount, bool declared) {
                if (_batchSink == null) return;
                // Without a declared dimension the columns are known only after the whole worksheet; such
                // worksheets are validated before the first row.
                if (!declared) throw new RowBatchDeclinedException();
                // The fast checks match element names only; rows are returned early only when they are SpreadsheetML.
                if (!WorksheetRootBindsSpreadsheetNamespace()) throw new RowBatchDeclinedException();
                // The batches held at once stay within the buffered-cell and chunk-row limits, so wide sheets flush sooner.
                int rows = Math.Min(RowBatchRows, _options.MaxDataReaderChunkRows);
                rows = (int)Math.Max(1L, Math.Min(rows, (long)_options.MaxDataReaderBufferedCells / RowBatchChannel.MaximumBatches / fieldCount));
                _batchSink.PublishRange(new RowBatchRange(firstRow, lastRow, firstColumn, fieldCount, declared), rows);
            }

            /// <summary>Copies one qualified row, whose metadata starts at <paramref name="rowOffset"/>, into the filling batch.</summary>
            private void AddRowToBatch(int rowIndex, int rowTagStart, int contentStart, int contentEnd, int rowOffset) {
                // Rows the compact indexer could not prove by itself, such as formulas with attributes, are returned once
                // the full XML check has read past them. The character check covers the row's own tag and attributes.
                if (!_sheetDataSupportsFastValidation || !RowTextIsValid(rowTagStart, contentEnd))
                    WaitForRowBatchXmlCheck(GetPartOffset(contentEnd));
                ThrowIfCellsOutsideDeclaredRange();
                _batchSink!.AddRow(this, rowIndex, contentStart, contentEnd, rowOffset);
            }

            private RowBatchXmlCheck? _batchXmlCheck;

            // After the character check finds text it cannot prove, such as malformed UTF-8, each row's own text is
            // checked; a row it cannot prove waits for the XML check like markup, and the rows before it still return.
#if NET8_0_OR_GREATER
            private bool RowTextIsValid(int start, int end) =>
                !_canonicalTextNeedsXmlReader || IsCanonicalDocumentTextValid(_buffer.AsSpan(start, end - start));
#else
            private static bool RowTextIsValid(int start, int end) => true;
#endif
            private CancellationToken _batchCt;

            /// <summary>Starts the full XML check of the worksheet on another thread, over its own read of the part.</summary>
            private RowBatchXmlCheck StartRowBatchXmlCheck() {
                OpenXmlPackagePartBufferReader? parts = _owner._partBufferReader?.TryOpenSibling();
                Stream? part = null;
                if (parts == null || !parts.TryOpenExactPartStream(_owner._worksheetPartName, _batchCt, out part)) {
                    parts?.Dispose();
                    // Without a second read of the part, the rest of the worksheet comes from the general reader.
                    throw new RowBatchHandoverException();
                }
                return new RowBatchXmlCheck(this, parts, part!, _batchCt);
            }

            private void WaitForRowBatchXmlCheck(long partOffset) {
                RowBatchXmlCheck check = _batchXmlCheck ??= StartRowBatchXmlCheck();
                check.WaitFor(partOffset, _batchCt);
                // The check found content that needs the SDK; rows from here come from the general reader.
                // A failed check is reported by WaitFor once a row reaches the failure.
                if (check.IsCompleted && !check.IsFaulted && !check.Result) throw new RowBatchHandoverException();
            }

            /// <summary>Waits for the full XML check at the end of the worksheet.</summary>
            private bool CompleteRowBatchXmlCheck(CancellationToken ct) {
                // Rows that all passed the fast checks need only the markup around them, as when validating first.
                if (_batchXmlCheck == null) {
                    using Stream? envelope = TryCreateWorksheetXmlEnvelope(ct);
                    if (envelope != null) return ValidateWorksheetXmlStream(envelope, ct);
                }
                RowBatchXmlCheck check = _batchXmlCheck ??= StartRowBatchXmlCheck();
                check.WaitFor(long.MaxValue, ct);
                return check.Result;
            }

            /// <summary>Runs the worksheet XML check over a separate read and reports how far it has read.</summary>
            private sealed class RowBatchXmlCheck : IDisposable {
                // The XML reader parses at most its internal buffers behind the bytes it has read.
                private const long ReadAheadSlack = 64 * 1024;
                private readonly Task<bool> _task;
                private readonly OpenXmlPackagePartBufferReader _parts;
                private long _read;

                internal RowBatchXmlCheck(ExcelUtf8RangeRowSource owner, OpenXmlPackagePartBufferReader parts, Stream part, CancellationToken ct) {
                    _parts = parts;
                    var counting = new CountingStream(part, this);
                    _task = Task.Run(() => {
                        try {
                            return owner.ValidateWorksheetXmlStream(counting, ct);
                        } finally {
                            counting.Dispose();
                        }
                    }, CancellationToken.None);
                }

                internal bool IsCompleted => _task.IsCompleted;

                internal bool IsFaulted => _task.IsFaulted;

                // Rethrows the check's own XML or validation error.
                internal bool Result => _task.GetAwaiter().GetResult();

                internal void WaitFor(long partOffset, CancellationToken ct) {
                    var spinner = new SpinWait();
                    while (!_task.IsCompleted && Volatile.Read(ref _read) - ReadAheadSlack < partOffset) {
                        ct.ThrowIfCancellationRequested();
                        if (!spinner.NextSpinWillYield) spinner.SpinOnce();
                        // A failed check is reported below through Result, with its own exception type.
                        else ((IAsyncResult)_task).AsyncWaitHandle.WaitOne(1);
                    }
                    if (!_task.IsCompleted) return;
                    // A failed check stops reading, so content further than the read-ahead behind it parsed cleanly:
                    // rows ending there are still returned, and the error comes from rows near the fault.
                    if (IsFaulted && Volatile.Read(ref _read) - ReadAheadSlack >= partOffset) return;
                    _ = Result;
                }

                public void Dispose() {
                    try { _task.Wait(); } catch { }
                    _parts.Dispose();
                }

                private sealed class CountingStream(Stream inner, RowBatchXmlCheck owner) : Stream {
                    public override int Read(byte[] buffer, int offset, int count) {
                        int read = inner.Read(buffer, offset, count);
                        Interlocked.Add(ref owner._read, read);
                        return read;
                    }
                    public override bool CanRead => true;
                    public override bool CanSeek => false;
                    public override bool CanWrite => false;
                    public override long Length => throw new NotSupportedException();
                    public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
                    public override void Flush() { }
                    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
                    public override void SetLength(long value) => throw new NotSupportedException();
                    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
                    protected override void Dispose(bool disposing) {
                        if (disposing) inner.Dispose();
                        base.Dispose(disposing);
                    }
                }
            }

            /// <summary>
            /// The reader's columns and rows come from the declared dimension. A cell outside it means the dimension is
            /// wrong; the rows before it were already returned, so the read that reaches it fails.
            /// </summary>
            private void ThrowIfCellsOutsideDeclaredRange() {
                if (_cellsFitWithinRange) return;
                throw new InvalidDataException(
                    $"Worksheet '{_owner._sheetName}' has cells outside its declared dimension after earlier rows were returned. " +
                    "Read it with ExcelReaderOptions.ValidateWorksheetBeforeFirstRow enabled to use the cells' actual range.");
            }

            private bool SelectBatchRow(int rowIndex, CancellationToken readCt) {
                while (true) {
                    RowBatch? batch = _currentBatch;
                    if (batch != null && _batchCursor < batch.Count) {
                        int physical = batch.RowIndexes[_batchCursor];
                        if (physical < rowIndex) {
                            _batchCursor++;
                            continue;
                        }
                        if (physical > rowIndex) return false;
                        _currentRowOffset = _batchCursor * _fieldCount;
                        _batchCursor++;
                        return true;
                    }
                    if (!TakeNextBatch(readCt)) {
                        RowBatchesEnded = true;
                        return false;
                    }
                }
            }

            private bool TakeNextBatch(CancellationToken readCt) {
                RowBatchChannel channel = _batchSource!;
                // The current row still refers to this batch until a replacement arrives; a canceled wait keeps it.
                RowBatch? next = channel.Take(readCt);
                if (_currentBatch != null) {
                    channel.Recycle(_currentBatch);
                    _currentBatch = null;
                }
                if (next == null) return false;
                _currentBatch = next;
                _batchCursor = 0;
                _buffer = next.Bytes;
                _length = next.ByteLength;
                _valueStarts = next.ValueStarts;
                _valueLengths = next.ValueLengths;
                _cellKinds = next.CellKinds;
                _formulaStarts = next.HasFormulas ? next.FormulaStarts : null;
                _formulaLengths = next.HasFormulas ? next.FormulaLengths : null;
#if NET8_0_OR_GREATER
                _canonicalTextHasNonSimpleBytes = next.HasNonSimpleText;
                _validatedCanonicalTextLength = next.ByteLength;
#endif
                if (_stringCache != null) Array.Clear(_stringCache, 0, _stringCache.Length);
                return true;
            }

            private bool TryGetRowBatchBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
                RowBatchRange range = _batchRange;
                firstRow = range.FirstRow;
                firstColumn = range.FirstColumn;
                lastRow = range.LastRow;
                lastColumn = range.FirstColumn + range.FieldCount - 1;
                return UsesRowBatches;
            }

            private void CompleteRowBatches(CancellationToken readCt) {
                // Rows after the last returned one may still fail validation; the last read reports it.
                while (TakeNextBatch(readCt)) { }
                RowBatchesEnded = true;
            }

            private void StopRowBatches() {
                if (_batchSource == null) return;
                // Batch arrays belong to the channel; detach them before the ordinary disposal returns arrays.
                _buffer = null;
                _valueStarts = null;
                _valueLengths = null;
                _cellKinds = null;
                _formulaStarts = null;
                _formulaLengths = null;
                _currentBatch = null;
                _batchStop!.Cancel();
                _batchSource.Cancel();
                try { _batchProducerTask!.GetAwaiter().GetResult(); } catch { }
                _batchProducer!.Dispose();
                _batchStop.Dispose();
                _batchSource.Release();
                _batchSource = null;
            }

            /// <summary>
            /// Raised to the reader after the rows the compact indexer qualified, when the rest of the worksheet needs
            /// the general XML reader.
            /// </summary>
            internal sealed class RowBatchHandoverException : Exception { }

            /// <summary>The worksheet is read with the validate-first reader; no row was returned.</summary>
            private sealed class RowBatchDeclinedException : Exception { }

            private readonly struct RowBatchRange {
                internal RowBatchRange(int firstRow, int lastRow, int firstColumn, int fieldCount, bool declared) {
                    FirstRow = firstRow;
                    LastRow = lastRow;
                    FirstColumn = firstColumn;
                    FieldCount = fieldCount;
                    Declared = declared;
                }
                internal int FirstRow { get; }
                internal int LastRow { get; }
                internal int FirstColumn { get; }
                internal int FieldCount { get; }
                internal bool Declared { get; }
            }

            /// <summary>One run of qualified rows; its arrays come from the shared pools and return on release.</summary>
            private sealed class RowBatch {
                internal byte[] Bytes = Array.Empty<byte>();
                internal int ByteLength;
                internal int Count;
                internal int[] RowIndexes = Array.Empty<int>();
                internal int[] ValueStarts = Array.Empty<int>();
                internal int[] ValueLengths = Array.Empty<int>();
                internal byte[] CellKinds = Array.Empty<byte>();
                internal int[] FormulaStarts = Array.Empty<int>();
                internal int[] FormulaLengths = Array.Empty<int>();
                internal bool HasFormulas;
                internal bool HasNonSimpleText;

                internal void Reset(int fieldCount, int rows) {
                    int cells = rows * fieldCount;
                    if (ValueStarts.Length < cells) {
                        ReturnCells();
                        ValueStarts = ArrayPool<int>.Shared.Rent(cells);
                        ValueLengths = ArrayPool<int>.Shared.Rent(cells);
                        CellKinds = ArrayPool<byte>.Shared.Rent(cells);
                    }
                    if (RowIndexes.Length < rows) {
                        if (RowIndexes.Length != 0) ArrayPool<int>.Shared.Return(RowIndexes);
                        RowIndexes = ArrayPool<int>.Shared.Rent(rows);
                    }
                    if (Bytes.Length == 0) Bytes = ArrayPool<byte>.Shared.Rent(1024 * 1024);
                    ByteLength = 0;
                    Count = 0;
                    HasFormulas = false;
                    HasNonSimpleText = false;
                }

                internal void EnsureBytes(int required) {
                    if (required <= Bytes.Length) return;
                    byte[] larger = ArrayPool<byte>.Shared.Rent(Math.Max(Bytes.Length * 2, required));
                    Buffer.BlockCopy(Bytes, 0, larger, 0, ByteLength);
                    ArrayPool<byte>.Shared.Return(Bytes, clearArray: true);
                    Bytes = larger;
                }

                internal void EnsureFormulas(int fieldCount, int rows) {
                    if (HasFormulas) return;
                    int cells = rows * fieldCount;
                    if (FormulaStarts.Length < cells) {
                        if (FormulaStarts.Length != 0) {
                            ArrayPool<int>.Shared.Return(FormulaStarts);
                            ArrayPool<int>.Shared.Return(FormulaLengths);
                        }
                        FormulaStarts = ArrayPool<int>.Shared.Rent(cells);
                        FormulaLengths = ArrayPool<int>.Shared.Rent(cells);
                    }
                    FormulaLengths.AsSpan(0, Count * fieldCount).Fill(-1);
                    HasFormulas = true;
                }

                private void ReturnCells() {
                    if (ValueStarts.Length == 0) return;
                    ArrayPool<int>.Shared.Return(ValueStarts);
                    ArrayPool<int>.Shared.Return(ValueLengths);
                    ArrayPool<byte>.Shared.Return(CellKinds);
                }

                internal void Return() {
                    ReturnCells();
                    if (RowIndexes.Length != 0) ArrayPool<int>.Shared.Return(RowIndexes);
                    // Row bytes hold workbook content, so they are cleared like other pooled part buffers.
                    if (Bytes.Length != 0) ArrayPool<byte>.Shared.Return(Bytes, clearArray: true);
                    if (FormulaStarts.Length != 0) {
                        ArrayPool<int>.Shared.Return(FormulaStarts);
                        ArrayPool<int>.Shared.Return(FormulaLengths);
                    }
                    Bytes = Array.Empty<byte>();
                    ValueStarts = ValueLengths = RowIndexes = FormulaStarts = FormulaLengths = Array.Empty<int>();
                    CellKinds = Array.Empty<byte>();
                }
            }

            /// <summary>Hands qualified rows from the background pass to the reader, at most three batches at a time.</summary>
            private sealed class RowBatchChannel {
                internal const int MaximumBatches = 3;
                private readonly object _gate = new();
                private readonly Queue<RowBatch> _ready = new();
                private readonly Stack<RowBatch> _free = new();
                private readonly List<RowBatch> _all = new(MaximumBatches);
                private int _created;
                private RowBatch? _filling;
                private RowBatchRange _range;
                private bool _rangePublished;
                private bool _declined;
                private bool _complete;
                private bool _canceled;
                private int _fieldCount;
                private int _batchRows;
                private ExceptionDispatchInfo? _fault;

                internal void PublishRange(RowBatchRange range, int batchRows) {
                    lock (_gate) {
                        _range = range;
                        _fieldCount = range.FieldCount;
                        _batchRows = batchRows;
                        _rangePublished = true;
                    }
                }

                /// <summary>Waits until the first batch is ready, so a worksheet that fails before any row still falls back.</summary>
                internal bool WaitForFirstRows(CancellationToken ct, out RowBatchRange range) {
                    lock (_gate) {
                        while (_ready.Count == 0 && !_complete && !_declined && _fault == null) {
                            ct.ThrowIfCancellationRequested();
                            Monitor.Wait(_gate, 20);
                        }
                        range = _range;
                        if (!_rangePublished || _declined) return false;
                        // A failure before any row was handed over leaves the decision to the ordinary readers.
                        if (_ready.Count == 0 && _fault != null) return false;
                        return true;
                    }
                }

                internal void AddRow(ExcelUtf8RangeRowSource producer, int rowIndex, int start, int end, int rowOffset) {
                    RowBatch batch = _filling ??= TakeFree();
                    int fields = _fieldCount;
                    int length = end - start;
                    batch.EnsureBytes(batch.ByteLength + length);
                    Buffer.BlockCopy(producer._buffer!, start, batch.Bytes, batch.ByteLength, length);
                    int delta = batch.ByteLength - start;
                    int target = batch.Count * fields;
                    int[] starts = producer._valueStarts!;
                    int[] lengths = producer._valueLengths!;
                    byte[] kinds = producer._cellKinds!;
                    for (int ordinal = 0; ordinal < fields; ordinal++) {
                        int source = rowOffset + ordinal;
                        int valueLength = lengths[source];
                        batch.ValueLengths[target + ordinal] = valueLength;
                        // Shared-string cells keep their table index; missing values keep their sentinel.
                        batch.ValueStarts[target + ordinal] = valueLength >= 0 ? starts[source] + delta : starts[source];
                        batch.CellKinds[target + ordinal] = kinds[source];
                    }
                    if (producer._formulaStarts != null) {
                        batch.EnsureFormulas(fields, _batchRows);
                        for (int ordinal = 0; ordinal < fields; ordinal++) {
                            int source = rowOffset + ordinal;
                            int formulaLength = producer._formulaLengths![source];
                            batch.FormulaLengths[target + ordinal] = formulaLength;
                            batch.FormulaStarts[target + ordinal] = formulaLength >= 0
                                ? producer._formulaStarts[source] + delta : producer._formulaStarts[source];
                        }
                    }
#if NET8_0_OR_GREATER
                    if (producer._canonicalTextHasNonSimpleBytes) batch.HasNonSimpleText = true;
#endif
                    batch.RowIndexes[batch.Count++] = rowIndex;
                    batch.ByteLength += length;
                    if (batch.Count == _batchRows) Flush();
                }

                internal void Flush() {
                    RowBatch? batch = _filling;
                    if (batch == null || batch.Count == 0) return;
                    _filling = null;
                    lock (_gate) {
                        _ready.Enqueue(batch);
                        Monitor.PulseAll(_gate);
                    }
                }

                private RowBatch TakeFree() {
                    lock (_gate) {
                        while (true) {
                            if (_canceled) throw new OperationCanceledException();
                            if (_free.Count > 0) {
                                RowBatch batch = _free.Pop();
                                batch.Reset(_fieldCount, _batchRows);
                                return batch;
                            }
                            if (_created < MaximumBatches) {
                                _created++;
                                var batch = new RowBatch();
                                _all.Add(batch);
                                batch.Reset(_fieldCount, _batchRows);
                                return batch;
                            }
                            Monitor.Wait(_gate, 20);
                        }
                    }
                }

                internal RowBatch? Take(CancellationToken ct) {
                    lock (_gate) {
                        while (_ready.Count == 0) {
                            if (_fault != null) _fault.Throw();
                            if (_complete || _declined) return null;
                            ct.ThrowIfCancellationRequested();
                            Monitor.Wait(_gate, 20);
                        }
                        return _ready.Dequeue();
                    }
                }

                internal void Recycle(RowBatch batch) {
                    lock (_gate) {
                        _free.Push(batch);
                        Monitor.PulseAll(_gate);
                    }
                }

                internal void Decline() {
                    lock (_gate) {
                        _declined = true;
                        Monitor.PulseAll(_gate);
                    }
                }

                internal void Complete() {
                    lock (_gate) {
                        _complete = true;
                        Monitor.PulseAll(_gate);
                    }
                }

                internal void Fault(Exception exception) {
                    lock (_gate) {
                        _fault ??= ExceptionDispatchInfo.Capture(exception);
                        Monitor.PulseAll(_gate);
                    }
                }

                internal void Cancel() {
                    lock (_gate) {
                        _canceled = true;
                        Monitor.PulseAll(_gate);
                    }
                }

                /// <summary>Returns every batch's arrays once the producer has stopped.</summary>
                internal void Release() {
                    lock (_gate) {
                        foreach (RowBatch batch in _all) batch.Return();
                        _all.Clear();
                        _ready.Clear();
                        _free.Clear();
                        _filling = null;
                    }
                }
            }
        }
    }
}
