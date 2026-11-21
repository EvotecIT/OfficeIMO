#nullable enable

using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const int StreamPublishPulseRows = 512;

            // Single-pass reading publishes indexed rows while inflation and indexing continue.
            // Validation of the remaining worksheet happens after earlier rows were delivered, so
            // failures surface at the read that reaches them or at the end of the sheet.
            private OpenXmlPackagePartBufferReader.StreamedWorksheetPart? _streamedPart;
            private Task? _streamIndexer;
            private CancellationTokenSource? _streamStop;
            private readonly object _streamGate = new object();
            private bool _streaming;
            private int _publishedRowCount;
            private bool _streamComplete;
            private ExceptionDispatchInfo? _streamFault;
            private int _streamFirstRow;
            private int _streamFirstColumn;
            private int _streamLastRow;
            private int _streamLastColumn;

            /// <summary>
            /// Starts single-pass reading when the caller opted in and the worksheet declares its
            /// dimension. Any other shape declines before a row is delivered.
            /// </summary>
            private static bool TryCreateStreamingUsedRange(ExcelSheetReader owner, CancellationToken ct,
                out ExcelUtf8RangeRowSource? source, out int declaredFirstColumn) {
                source = null;
                declaredFirstColumn = 0;
                if (!owner._opt.EnableSinglePassReading
                    || owner._partBufferReader?.AllowsPipelining != true
                    || !owner._partBufferReader.TryStartStreamedRead(owner._worksheetPartName, MaximumBufferSize, ct,
                        out OpenXmlPackagePartBufferReader.StreamedWorksheetPart? part)) return false;

                var candidate = new ExcelUtf8RangeRowSource(owner, part!.Buffer, part.Length) {
                    _ownsBuffer = false,
                    _inflation = part.Progress,
                    _expectedLength = part.Length,
                    _streamedPart = part
                };
                try {
                    // The dimension precedes sheet data; wait for a bounded prefix before looking for it.
                    int wanted = Math.Min(part.Length, 64 * 1024);
                    candidate._length = part.Progress.WaitForMore(0, ct);
                    while (candidate._length < wanted) {
                        int next = part.Progress.WaitForMore(candidate._length, ct);
                        if (next <= candidate._length) break;
                        candidate._length = next;
                    }
                    if (!candidate.TryGetDeclaredRange(ct, out int firstRow, out int firstColumn,
                            out int lastRow, out int lastColumn)) {
                        candidate.Dispose();
                        return false;
                    }
                    int fields = lastColumn - firstColumn + 1;
                    long rows = (long)lastRow - firstRow + 1;
                    if (firstRow <= 0 || lastRow < firstRow || lastRow > ExcelA1.MaxRows
                        || firstColumn <= 0 || lastColumn < firstColumn || lastColumn > ExcelA1.MaxColumns
                        || fields > owner._opt.MaxDataReaderColumns || owner._opt.MaxDataReaderChunkRows <= 0
                        || rows * fields > Math.Min(MaximumIndexedCells, owner._opt.MaxDataReaderBufferedCells)) {
                        candidate.Dispose();
                        return false;
                    }

                    candidate.InitializeRange(firstRow, lastRow, firstColumn, fields);
                    // Metadata arrays never move while the reader holds published rows.
                    candidate.EnsureRowCapacity((int)rows);
#if NET8_0_OR_GREATER
                    candidate.EnsureDecodedValues();
                    Array.Clear(candidate._decodedTags!, 0, candidate._decodedTags!.Length);
#endif
                    candidate._streaming = true;
                    candidate._streamFirstRow = firstRow;
                    candidate._streamFirstColumn = firstColumn;
                    candidate._streamLastRow = lastRow;
                    candidate._streamLastColumn = lastColumn;
                    candidate._streamStop = CancellationTokenSource.CreateLinkedTokenSource(ct);
                    CancellationToken token = candidate._streamStop.Token;
                    candidate._streamIndexer = Task.Run(() => candidate.RunStreamIndexer(firstRow, lastRow, token),
                        CancellationToken.None);
                    declaredFirstColumn = firstColumn;
                    source = candidate;
                    return true;
                } catch {
                    candidate.Dispose();
                    throw;
                }
            }

            private void RunStreamIndexer(int firstRow, int lastRow, CancellationToken ct) {
                try {
                    bool qualified = TryIndexRows(firstRow, lastRow, ct);
                    _length = _inflation!.WaitForCompletion(ct);
                    Exception? producerFault = _streamedPart!.Fault;
                    if (producerFault != null) ExceptionDispatchInfo.Capture(producerFault).Throw();
                    if (qualified) {
                        qualified = _cellsFitWithinRange
                            && (_validatedIndexedXml || IsCanonicalWorksheetXmlFullyValidated(ct) || ValidateBufferedWorksheetXml(ct));
                    }
                    if (!qualified) {
                        // Malformed XML raises the exception the default reader's XML check raises.
                        ValidateBufferedWorksheetXml(ct);
                        throw new NotSupportedException(
                            $"Worksheet '{_owner._sheetName}' contains content that single-pass reading cannot index after rows were returned. Read it with EnableSinglePassReading disabled.");
                    }
                } catch (Exception exception) {
                    _streamFault = ExceptionDispatchInfo.Capture(exception);
                } finally {
                    lock (_streamGate) {
                        Volatile.Write(ref _publishedRowCount, _rowCount);
                        Volatile.Write(ref _streamComplete, true);
                        Monitor.PulseAll(_streamGate);
                    }
                }
            }

            /// <summary>Makes rows indexed so far visible to the reader thread.</summary>
            private void PublishStreamRows() {
                int previous = _publishedRowCount;
                Volatile.Write(ref _publishedRowCount, _rowCount);
                if (previous / StreamPublishPulseRows != _rowCount / StreamPublishPulseRows) {
                    lock (_streamGate) Monitor.PulseAll(_streamGate);
                }
            }

            /// <summary>Waits until the requested row, a later row or the end of indexing is published.</summary>
            private int WaitForStreamRow(int rowIndex, CancellationToken readCt) {
                var spinner = new SpinWait();
                while (true) {
                    int published = Volatile.Read(ref _publishedRowCount);
                    int cursor = _rowCursor;
                    while (cursor < published && _rowIndexes![cursor] < rowIndex) cursor++;
                    if (cursor < published) return published;
                    if (Volatile.Read(ref _streamComplete)) {
                        _streamFault?.Throw();
                        return Volatile.Read(ref _publishedRowCount);
                    }
                    readCt.ThrowIfCancellationRequested();
                    if (!spinner.NextSpinWillYield) {
                        spinner.SpinOnce();
                        continue;
                    }
                    lock (_streamGate) {
                        if (Volatile.Read(ref _publishedRowCount) == published && !Volatile.Read(ref _streamComplete)) {
                            Monitor.Wait(_streamGate, 1);
                        }
                    }
                    spinner = new SpinWait();
                }
            }

            /// <summary>Waits for the remaining worksheet to be validated at the end of the rows.</summary>
            internal void CompleteStreaming(CancellationToken readCt) {
                if (!_streaming) return;
                while (!Volatile.Read(ref _streamComplete)) {
                    readCt.ThrowIfCancellationRequested();
                    lock (_streamGate) {
                        if (!Volatile.Read(ref _streamComplete)) Monitor.Wait(_streamGate, 1);
                    }
                }
                _streamFault?.Throw();
            }

            private bool TryGetStreamingBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
                firstRow = _streamFirstRow;
                firstColumn = _streamFirstColumn;
                lastRow = _streamLastRow;
                lastColumn = _streamLastColumn;
                return _streaming;
            }

            /// <summary>Stops background indexing and inflation before metadata or the buffer is released.</summary>
            private void StopStreaming() {
                if (_streamIndexer != null) {
                    _streamStop!.Cancel();
                    try { _streamIndexer.GetAwaiter().GetResult(); } catch (Exception) { }
                    _streamIndexer = null;
                }
                _streamStop?.Dispose();
                _streamStop = null;
                _streamedPart?.Dispose();
                _streamedPart = null;
            }
        }
    }
}
