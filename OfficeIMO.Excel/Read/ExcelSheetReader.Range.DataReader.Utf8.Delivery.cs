#nullable enable

using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelXmlRangeDataReader {
            /// <summary>
            /// Indexed rows pass their read token directly to window selection. They
            /// do not run XmlReader callbacks, so no temporary XML cancellation state
            /// or finally restoration is needed. Existing rows survive a canceled refill.
            /// </summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool ReadIndexedCore(CancellationToken cancellationToken) {
                _ct.ThrowIfCancellationRequested();
                cancellationToken.ThrowIfCancellationRequested();
                if (_closed) return false;
                ThrowIfXmlTextBufferingFailed();
                if (TryReadIndexedLogicalRow(cancellationToken, out _)) return true;
                ReleaseCurrentValueReferences();
                _currentRow = null;
                return false;
            }

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private bool TryReadIndexedLogicalRow(CancellationToken cancellationToken, out object?[] row) {
                row = Array.Empty<object?>();
                bool hasPhysicalRow;
                try {
                    if (_nextLogicalRow > _lastRow) {
                        _utf8Source!.CompleteStreaming(cancellationToken);
                        return false;
                    }
                    hasPhysicalRow = _utf8Source!.SelectRow(_nextLogicalRow, _ct, cancellationToken);
                    if (!hasPhysicalRow && _utf8Source.RowBatchesEnded) {
                        // Without a declared dimension the last row is known only when the background pass ends.
                        _utf8Source.CompleteStreaming(cancellationToken);
                        return false;
                    }
                } catch (ExcelUtf8RangeRowSource.RowBatchHandoverException) {
                    _utf8Source!.Dispose();
                    _utf8Source = null;
                    _handoverPending = true;
                    // The XML reader replays the worksheet up to the next row; it observes this read's token.
                    _activeReadCancellationToken = cancellationToken;
                    try {
                        CompleteHandover(cancellationToken);
                        if (!TryReadLogicalRow(out row)) return false;
                    } finally {
                        _activeReadCancellationToken = _ct;
                    }
                    _currentRow = row;
                    return true;
                }
                ReleaseCurrentValueReferences();
                row = hasPhysicalRow ? _currentValues : _blankRow;
                _currentRow = row;
                _currentRowIsBlank = !hasPhysicalRow;
                _currentRowActive = hasPhysicalRow;
                _currentRowFinished = !hasPhysicalRow;
                _nextLogicalRow++;
                return true;
            }

            /// <summary>
            /// Continues a worksheet whose remaining content the compact indexer does not handle with the general XML
            /// reader, from the next logical row. The whole part is checked first, as the reader would have done before
            /// the first row; rows already returned keep their columns. A canceled read leaves the switch to the next
            /// read; any other failure is reported by every later read.
            /// </summary>
            private void CompleteHandover(CancellationToken cancellationToken) {
                if (!_handoverPending) return;
                try {
                    StartXmlRowsAfterHandover(cancellationToken);
                } catch (OperationCanceledException) when (!_ct.IsCancellationRequested) {
                    throw;
                } catch (Exception exception) {
                    _handoverPending = false;
                    _handoverFailure = ExceptionDispatchInfo.Capture(exception);
                    throw;
                }
                _handoverPending = false;
            }

            private bool _handoverPending;
            private ExceptionDispatchInfo? _handoverFailure;

            private void ThrowIfHandoverFailed() => _handoverFailure?.Throw();

            private void StartXmlRowsAfterHandover(CancellationToken cancellationToken) {
                // The indexed row is complete; XML mode must not skip content for it.
                ReleaseCurrentValueReferences();
                _currentRow = null;
                _currentRowActive = false;
                _currentRowFinished = true;
                _currentRowIsBlank = false;

                // The check records the actual cell bounds and row order only when neither is cached yet.
                string? cachedRange = _owner._usedRangeA1;
                _owner._usedRangeA1 = null;
                Volatile.Write(ref _owner._implicitXmlRowIndexes, null);
                bool rowsStrictlyIncreasing;
                string? cellRange;
                try {
                    using (Stream validation = _owner.OpenReaderWorksheetStream(cancellationToken)) {
                        if (!TryPrepareWorksheetStream(validation)) {
                            throw new InvalidDataException($"Worksheet '{_owner._sheetName}' could not be reopened for reading.");
                        }
                        rowsStrictlyIncreasing = _owner.ValidateDataReaderProjectionXml(validation, cancellationToken);
                    }
                    cellRange = _owner._usedRangeA1;
                } finally {
                    _owner._usedRangeA1 = cachedRange ?? _owner._usedRangeA1;
                }
                // Rows were returned, so the worksheet has cells. No bounds means the check could not collect them,
                // as for worksheets with tables, and proved no row order either.
                if (cellRange == null)
                    rowsStrictlyIncreasing = _owner.RowsAreSortedWithinRangeXmlFast(_firstRow, _lastRow, cancellationToken);
                if (!rowsStrictlyIncreasing) {
                    // Rows before the next one were already returned in file order.
                    throw new InvalidDataException(
                        $"Worksheet '{_owner._sheetName}' lists row {_nextLogicalRow} or later out of order after earlier rows were returned. " +
                        "Read it with ExcelReaderOptions.ValidateWorksheetBeforeFirstRow enabled.");
                }
                if (cellRange != null) {
                    var (_, usedFirstColumn, usedLastRow, usedLastColumn) = ExcelA1.ParseRange(cellRange);
                    if (usedFirstColumn < _firstColumn || usedLastColumn > _lastColumn || usedLastRow > _lastRow) {
                        throw new InvalidDataException(
                            $"Worksheet '{_owner._sheetName}' has cells outside its declared dimension " +
                            "after rows inside it were returned. Read it with ExcelReaderOptions.ValidateWorksheetBeforeFirstRow enabled.");
                    }
                    _lastRow = usedLastRow;
                }
                Stream stream = _owner.OpenReaderWorksheetStream(cancellationToken);
                try {
                    RewindWorksheetStream(stream);
                    _reader = OpenWorksheetXmlReader(stream);
                } catch {
                    stream.Dispose();
                    throw;
                }
                _stream = stream;
                _xmlTextBudget = new XmlDataReaderTextBudget(_owner._opt.MaxXmlDataReaderBufferedCharacters,
                    ThrowIfReadCancellationRequested);
                // Indexed rows report presence from their cell index; XML rows track it as the constructor does.
                if (_trackCellPresence) _currentCellPresence ??= new bool[_fieldCount];
            }
        }
    }
}
