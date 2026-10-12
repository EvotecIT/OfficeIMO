#nullable enable

using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Commits the provisional next row only after every cell has been proven
            /// to omit its coordinate. A later explicit cell reference may determine
            /// the row number, so any incomplete attempt restores the discovery state
            /// before the ordinary coordinate lookahead runs.
            /// A committed row initializes every metadata slot; callers reset partial
            /// metadata before indexing an unsuccessful attempt through the fallback.
            /// </summary>
            private bool TryIndexCanonicalImplicitDenseRow(
                ref int position,
                int rowIndex,
                int rowOffset,
                CancellationToken ct) {
                int originalPosition = position;
                int minimumCellRow = _minimumCellRow;
                int maximumCellRow = _maximumCellRow;
                int minimumCellColumn = _minimumCellColumn;
                int maximumCellColumn = _maximumCellColumn;
                if (TryIndexCanonicalImplicitDenseRowCore(ref position, rowIndex, rowOffset, ct)) {
                    return true;
                }

                position = originalPosition;
                _minimumCellRow = minimumCellRow;
                _maximumCellRow = maximumCellRow;
                _minimumCellColumn = minimumCellColumn;
                _maximumCellColumn = maximumCellColumn;
                return false;
            }

            private bool TryIndexCanonicalImplicitDenseRowCore(
                ref int position,
                int rowIndex,
                int rowOffset,
                CancellationToken ct) {
                int columnIndex = 1;
                int cellsUntilCancellationCheck = 0;
                Utf8ImplicitCellTagCacheEntry[]? tagCache = _implicitCellTagCache;
                while (position < _length) {
                    if (position <= _length - 6
                        && _buffer![position + 1] == (byte)'/'
                        && MatchesUtf8(position, "</row>"u8)) {
                        // Every indexed cell owns its metadata. Only absent trailing
                        // cells need missing sentinels when a complete row is proven.
                        int indexedFieldCount = columnIndex - 1;
                        if (indexedFieldCount < _fieldCount) {
                            InitializeMetadataRow(rowOffset, indexedFieldCount);
                        }
                        position += 6;
                        UpdateUsedBounds(rowIndex, columnIndex - 1);
                        return true;
                    }

                    if (cellsUntilCancellationCheck-- == 0) {
                        ct.ThrowIfCancellationRequested();
                        cellsUntilCancellationCheck = 256;
                    }
                    if (columnIndex == 1 && _minimumCellColumn > 1) {
                        _minimumCellColumn = 1;
                    }
                    int ordinal = columnIndex - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        return false;
                    }

                    tagCache ??= RentImplicitCellTagCache();
                    ref Utf8ImplicitCellTagCacheEntry cachedTag = ref tagCache[ordinal];
                    int cellStart = position;
                    bool dateStylesEnabled = TreatDatesForIndex;
                    bool repeatedTag = cachedTag.Length > 0
                        && cachedTag.DateStylesEnabled == dateStylesEnabled
                        && position <= _length - cachedTag.Length
                        && _buffer!.AsSpan(position, cachedTag.Length).SequenceEqual(
                            _buffer.AsSpan(cachedTag.Start, cachedTag.Length));
                    Utf8CellKind kind;
                    int styleIndex;
                    bool isEmpty;
                    if (repeatedTag) {
                        kind = cachedTag.Kind;
                        styleIndex = cachedTag.StyleIndex;
                        isEmpty = cachedTag.IsEmpty;
                        position += cachedTag.Length;
                    } else if (!TryReadCanonicalImplicitCellAttributes(
                                   ref position, out kind, out styleIndex, out isEmpty)) {
                        return false;
                    }
                    int tagLength = position - cellStart;

                    int cellIndex = rowOffset + ordinal;
                    bool hasCachedValue = false;
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!isEmpty
                        && !TryIndexCompactValueCell(
                            ref position, cellIndex, kind,
                            out hasCachedValue, out valueStart, out valueLength)) {
                        return false;
                    }

                    int sharedStringIndex = ValidateIndexedCell(
                        rowIndex, columnIndex, kind, styleIndex,
                        sharedFormulaFollower: false, hasCachedValue, valueStart, valueLength);
                    if (sharedStringIndex >= 0) {
                        _valueStarts![cellIndex] = sharedStringIndex;
                        _valueLengths![cellIndex] = SharedStringIndexValueLength;
                    } else if (valueLength < 0) {
                        _valueStarts![cellIndex] = 0;
                        _valueLengths![cellIndex] = -1;
                    }
                    if (_formulaStarts != null) {
                        _formulaStarts[cellIndex] = 0;
                        _formulaLengths![cellIndex] = -1;
                    }
                    byte encodedKind = repeatedTag ? cachedTag.EncodedKind : EncodeCellKind(kind, styleIndex);
                    _cellKinds![cellIndex] = encodedKind;
                    if (!repeatedTag) {
                        cachedTag = new Utf8ImplicitCellTagCacheEntry(
                            cellStart, tagLength, kind, styleIndex, isEmpty, encodedKind, dateStylesEnabled);
                    }
                    columnIndex++;
                }

                return false;
            }
        }
    }
}
