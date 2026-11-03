#nullable enable

using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private bool _ownsBuffer = true;
            private bool _unknownRowExtent;
            private WorksheetInflationProgress? _inflation;

            private static bool TryCreatePipelinedUsedRange(
                ExcelSheetReader owner, CancellationToken ct,
                out ExcelUtf8RangeRowSource? source, out int firstColumn,
                out byte[]? buffer, out int length) {
                source = null;
                firstColumn = 0;
                buffer = null;
                length = 0;
                if (owner._partBufferReader?.AllowsPipelining != true) return false;
                ExcelUtf8RangeRowSource? candidate = null;
                int column = 0;
                try {
                    bool attempted = owner._partBufferReader.TryReadPipelined(
                        owner._worksheetPartName, MaximumBufferSize, ct,
                        (bytes, total, progress) => {
                            candidate = new ExcelUtf8RangeRowSource(owner, bytes, total) {
                                _ownsBuffer = false,
                                _inflation = progress
                            };
                            candidate._length = progress.WaitForMore(0, ct);
                            candidate.ValidateProducedCanonicalText(ct);
                            return candidate.TryPrepareUsedRange(ct, allowUnknownRows: true, out column);
                        }, out buffer, out length, out bool qualified);
                    if (!attempted) return false;
                    if (qualified) {
                        candidate!._ownsBuffer = true;
                        candidate._inflation = null;
                        source = candidate;
                        candidate = null;
                        firstColumn = column;
                        buffer = null;
                    }
                    return true;
                } finally {
                    candidate?.Dispose();
                }
            }

            /// <summary>
            /// Qualifies the whole immutable worksheet before publication. A pipelined
            /// first-row proposal uses a bounded unknown extent; actual coordinates
            /// and complete XML proof still come from the full indexing pass.
            /// </summary>
            private bool TryPrepareUsedRange(CancellationToken ct, bool allowUnknownRows, out int declaredFirstColumn) {
                declaredFirstColumn = 0;
                bool declared = TryGetDeclaredRange(ct,
                    out int firstRow, out int firstColumn, out int lastRow, out int lastColumn);
                _unknownRowExtent = !declared && allowUnknownRows;
                if (!declared && !TryInferUndeclaredRange(ct,
                    out firstRow, out firstColumn, out lastRow, out lastColumn, _unknownRowExtent)) return false;
                int fields = lastColumn - firstColumn + 1;
                if (firstRow <= 0 || lastRow < firstRow || lastRow > ExcelA1.MaxRows
                    || firstColumn <= 0 || lastColumn < firstColumn || lastColumn > ExcelA1.MaxColumns
                    || fields <= 0 || fields > _options.MaxDataReaderColumns
                    || fields > _options.MaxDataReaderBufferedCells || _options.MaxDataReaderChunkRows <= 0) return false;
                InitializeRange(firstRow, lastRow, firstColumn, fields);
                if (!TryIndexRows(firstRow, lastRow, ct)) return false;
                // A row-like token in trailing comments may publish sheetData's end
                // before the root XML is complete. Global proof needs the full part.
                if (_inflation != null) _length = _inflation.WaitForCompletion(ct);
                if (!IsCanonicalWorksheetXmlFullyValidated(ct) && !ValidateBufferedWorksheetXml(ct)) return false;
                ct.ThrowIfCancellationRequested();
                declaredFirstColumn = firstColumn;
                return true;
            }

            private bool TryAdvanceInflation(CancellationToken ct) {
                if (_inflation == null) return false;
                int available = _inflation.WaitForMore(_length, ct);
                if (available <= _length) return false;
                _length = available;
                ValidateProducedCanonicalText(ct);
                return true;
            }

            /// <summary>
            /// Promotes an unknown physical row count to the existing rolling index
            /// before its metadata exceeds the cell budget. Qualified row positions
            /// survive; discarded cell metadata is rebuilt by ordinary window reads.
            /// </summary>
            private void PromoteProvisionalIndexToWindow() {
                long maximumCells = Math.Min(MaximumIndexedCells, _options.MaxDataReaderBufferedCells);
                _cellWindowRowCapacity = (int)Math.Min(_options.MaxDataReaderChunkRows, maximumCells / _fieldCount);
                _indexedTreatDatesUsingNumberFormat = _options.TreatDatesUsingNumberFormat;
                _indexedUseCachedFormulaResult = _options.UseCachedFormulaResult;
                int cells = checked(_cellWindowRowCapacity * _fieldCount);
                bool hadFormulas = _formulaStarts != null;
                ReturnRowArray(ref _valueStarts);
                ReturnRowArray(ref _valueLengths);
                ReturnRowArray(ref _formulaStarts);
                ReturnRowArray(ref _formulaLengths);
                if (_cellKinds != null) {
                    byte[] previous = _cellKinds;
                    _cellKinds = null;
                    ArrayPool<byte>.Shared.Return(previous);
                }
                _valueStarts = ArrayPool<int>.Shared.Rent(cells);
                _valueLengths = ArrayPool<int>.Shared.Rent(cells);
                _cellKinds = ArrayPool<byte>.Shared.Rent(cells);
                // Qualification must retain formula-array presence, even if all
                // formula cells occurred before promotion. Window replay relies on
                // those arrays already existing before it initializes earlier cells.
                if (hadFormulas) EnsureFormulaMetadata();
            }
        }
    }
}
