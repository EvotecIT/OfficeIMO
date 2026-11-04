#nullable enable

using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private bool _cachedImplicitRowReady;
            private bool _cachedImplicitRowRejected;

            /// <summary>
            /// Reuses a completely qualified implicit layout without carrying tag
            /// discovery and cache updates through its hot loop. Exact tag matches
            /// retain style/type proof; values and shared-string indexes are per-cell.
            /// A miss leaves used bounds untouched so ordinary row discovery can retry.
            /// </summary>
            private bool TryIndexCachedImplicitRow(ref int position, int rowIndex, int rowOffset, CancellationToken ct) {
                Utf8ImplicitCellTagCacheEntry[] tags = _implicitCellTagCache!;
                int cellsUntilCancellationCheck = 0;
                for (int ordinal = 0; ordinal < _fieldCount; ordinal++) {
                    if (cellsUntilCancellationCheck-- == 0) {
                        ct.ThrowIfCancellationRequested();
                        cellsUntilCancellationCheck = 256;
                    }
                    ref Utf8ImplicitCellTagCacheEntry tag = ref tags[ordinal];
                    if (tag.Length <= 0 || tag.DateStylesEnabled != TreatDatesForIndex
                        || position > _length - tag.Length
                        || !_buffer!.AsSpan(position, tag.Length).SequenceEqual(_buffer.AsSpan(tag.Start, tag.Length))) {
                        return false;
                    }
                    position += tag.Length;
                    int cellIndex = rowOffset + ordinal;
                    bool hasCachedValue = false;
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!tag.IsEmpty && !TryIndexCompactValueCell(ref position, cellIndex, tag.Kind,
                            out hasCachedValue, out valueStart, out valueLength)) {
                        return false;
                    }
                    int sharedStringIndex = tag.Kind == Utf8CellKind.SharedString
                        ? ValidateIndexedCell(rowIndex, ordinal + 1, tag.Kind, tag.StyleIndex,
                            sharedFormulaFollower: false, hasCachedValue, valueStart, valueLength)
                        : -1;
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
                    _cellKinds![cellIndex] = tag.EncodedKind;
                }
                if (!MatchesUtf8(position, "</row>"u8)) return false;
                position += 6;
                if (_minimumCellColumn > 1) _minimumCellColumn = 1;
                UpdateUsedBounds(rowIndex, _fieldCount);
                return true;
            }
        }
    }
}
