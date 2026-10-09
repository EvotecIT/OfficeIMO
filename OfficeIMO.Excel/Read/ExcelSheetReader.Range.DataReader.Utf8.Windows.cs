#nullable enable

using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private int _cellWindowRowCapacity;
            private int _cellWindowFirstRow = -1;
            private int _cellWindowRowCount;
            private int _selectedPhysicalRow = -1;
            private int[]? _rowContentStarts;
            private int[]? _rowContentEnds;
            private bool _indexedTreatDatesUsingNumberFormat;
            private bool _indexedUseCachedFormulaResult;

            private bool UsesCellWindow => _cellWindowRowCapacity > 0;

            // Whole-sheet indexes fix these decisions during opening. A rolling
            // index must retain the same decisions when later windows are filled.
            private bool TreatDatesForIndex => UsesCellWindow
                ? _indexedTreatDatesUsingNumberFormat : _options.TreatDatesUsingNumberFormat;

            private bool UseCachedFormulaResultForIndex => UsesCellWindow
                ? _indexedUseCachedFormulaResult : _options.UseCachedFormulaResult;

            /// <summary>
            /// Keeps cell metadata within the existing per-reader buffering and
            /// chunk-row limits. The complete qualification pass reuses one row;
            /// only physical row coordinates and byte positions survive that pass.
            /// </summary>
            private int InitializeCellWindow(int rowCapacity, long rangeRows, int fieldCount) {
                long maximumCells = Math.Min(MaximumIndexedCells, _options.MaxDataReaderBufferedCells);
                if (rangeRows * fieldCount <= maximumCells) {
                    return rowCapacity;
                }

                _cellWindowRowCapacity = (int)Math.Min(
                    _options.MaxDataReaderChunkRows, maximumCells / fieldCount);
                _indexedTreatDatesUsingNumberFormat = _options.TreatDatesUsingNumberFormat;
                _indexedUseCachedFormulaResult = _options.UseCachedFormulaResult;
                _rowContentStarts = ArrayPool<int>.Shared.Rent(rowCapacity);
                _rowContentEnds = ArrayPool<int>.Shared.Rent(rowCapacity);
                return _cellWindowRowCapacity;
            }

            private int GetQualificationRowOffset(int physicalRow) => UsesCellWindow
                ? 0 : checked(physicalRow * _fieldCount);

            private void RetainQualifiedRowPosition(int physicalRow, int contentStart, int contentEnd) {
                if (UsesCellWindow) {
                    _rowContentStarts![physicalRow] = contentStart;
                    _rowContentEnds![physicalRow] = contentEnd;
                }
            }

            private void EnsureWindowRowCapacity(int required) {
                int currentCapacity = Math.Min(_rowIndexes!.Length,
                    Math.Min(_rowContentStarts!.Length, _rowContentEnds!.Length));
                if (required <= currentCapacity) return;

                int nextCapacity = Math.Min(A1.MaxRows, checked(currentCapacity * 2));
                GrowRowArray(ref _rowIndexes, nextCapacity, _rowCount);
                GrowRowArray(ref _rowContentStarts, nextCapacity, _rowCount);
                GrowRowArray(ref _rowContentEnds, nextCapacity, _rowCount);
            }

            private int GetInitializedFormulaCellCount() => UsesCellWindow
                ? _fieldCount : checked((_rowCount + 1) * _fieldCount);

            /// <summary>
            /// Reuses the ordinary cell indexers over previously qualified immutable
            /// rows. A canceled refill restores the selected row, so its unloaded
            /// getters and borrowed text remain usable before reading is retried.
            /// </summary>
            private int SelectCellWindowRow(
                int physicalRow, CancellationToken lifetimeCt, CancellationToken readCt) {
                if (physicalRow < _cellWindowFirstRow
                    || physicalRow >= _cellWindowFirstRow + _cellWindowRowCount) {
                    _cellWindowFirstRow = -1;
                    _cellWindowRowCount = 0;
                    int count = Math.Min(_cellWindowRowCapacity, _rowCount - physicalRow);
                    CancellationTokenSource? linkedCancellation = lifetimeCt.CanBeCanceled
                        && readCt.CanBeCanceled && lifetimeCt != readCt
                            ? CancellationTokenSource.CreateLinkedTokenSource(lifetimeCt, readCt) : null;
                    CancellationToken ct = linkedCancellation?.Token
                        ?? (readCt.CanBeCanceled ? readCt : lifetimeCt);
                    int qualifiedLastCellRow = _maximumCellRow;
                    try {
                        for (int index = 0; index < count; index++) {
                            ct.ThrowIfCancellationRequested();
                            IndexQualifiedWindowRow(physicalRow + index, checked(index * _fieldCount), ct);
                        }
                        ct.ThrowIfCancellationRequested();
                    } catch (OperationCanceledException) {
                        if (_selectedPhysicalRow >= 0) {
                            // Restoration is bounded to one already validated row.
                            IndexQualifiedWindowRow(_selectedPhysicalRow, 0, CancellationToken.None);
                            _currentRowOffset = 0;
                            _cellWindowFirstRow = _selectedPhysicalRow;
                            _cellWindowRowCount = 1;
                        }
                        throw;
                    } finally {
                        // The normal indexers advance the discovered last row.
                        // A refill revisits earlier rows after full qualification.
                        _maximumCellRow = qualifiedLastCellRow;
                        linkedCancellation?.Dispose();
                    }
                    _cellWindowFirstRow = physicalRow;
                    _cellWindowRowCount = count;
                }
                return checked((physicalRow - _cellWindowFirstRow) * _fieldCount);
            }

            private void IndexQualifiedWindowRow(int physicalRow, int rowOffset, CancellationToken ct) {
                InitializeMetadataRow(rowOffset);
                int contentStart = _rowContentStarts![physicalRow];
                int contentEnd = _rowContentEnds![physicalRow];
                if (contentStart == contentEnd) return; // An empty row start tag.

                int position = contentStart;
                int rowIndex = _rowIndexes![physicalRow];
                bool indexed = TryIndexCanonicalImplicitDenseRow(ref position, rowIndex, rowOffset, ct);
                if (!indexed) {
                    position = contentStart;
                    InitializeMetadataRow(rowOffset);
                    indexed = TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
                }
                if (!indexed) {
                    position = contentStart;
                    InitializeMetadataRow(rowOffset);
                    indexed = TryIndexRow(ref position, rowIndex, rowOffset, rowWithinRange: true, ct);
                }
                if (!indexed || position != contentEnd) {
                    throw new InvalidDataException("A qualified worksheet row could not be indexed.");
                }
            }
        }
    }
}
