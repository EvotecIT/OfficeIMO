using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelSheetReader {
    private sealed partial class ExcelUtf8RangeRowSource {
        private bool _mappingRowsQualified;
        private bool _mappingBufferFromSharedPool;

        internal bool CanCopyMappingRows => !_disposed && !UsesCellWindow && _mappingRowsQualified;

        internal void AdvanceMappingRows(int nextRow) {
            while (_rowCursor < _rowCount && _rowIndexes![_rowCursor] < nextRow) _rowCursor++;
        }

        /// <summary>Copies qualified scalar payloads into independently owned worker storage.</summary>
        internal bool TryCopyMappingRows(ExcelSheetReader owner, int firstRow, int count,
            int firstOrdinal, int fieldCount, CancellationToken ct, out ExcelUtf8RangeRowSource? copy) {
            copy = null;
            int firstPhysical = _rowCursor;
            while (firstPhysical < _rowCount && _rowIndexes![firstPhysical] < firstRow) firstPhysical++;
            int physical = firstPhysical, minimum = _length, maximum = 0;
            int savedOffset = _currentRowOffset;
            try {
                for (int row = 0; row < count; row++) {
                    ct.ThrowIfCancellationRequested();
                    if (physical >= _rowCount || _rowIndexes![physical] != firstRow + row) continue;
                    int offset = physical * _fieldCount + firstOrdinal;
                    _currentRowOffset = physical * _fieldCount;
                    for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                        int cell = offset + ordinal;
                        var kind = (Utf8CellKind)(_cellKinds![cell] & CellKindMask);
                        if (kind is Utf8CellKind.SharedString or Utf8CellKind.Error ||
                            (_formulaLengths is not null && _formulaLengths[cell] >= 0)) return false;
                        int length = _valueLengths![cell];
                        if (length < 0) continue;
                        if ((_cellKinds[cell] & DateStyleCellKindFlag) != 0 &&
                            !TryGetDateTime(firstOrdinal + ordinal, out _)) return false;
                        minimum = Math.Min(minimum, _valueStarts![cell]);
                        maximum = Math.Max(maximum, _valueStarts![cell] + length);
                    }
                    physical++;
                }
            } catch (ArgumentException) { return false; }
            finally { _currentRowOffset = savedOffset; }
            if (TryShareMappingRows(owner, firstRow, count, firstPhysical,
                firstOrdinal, fieldCount, ct, out copy)) return true;
            int lengthToCopy = maximum > minimum ? maximum - minimum : 0;
            // Worker batches need concurrent leases of the same size.
            byte[] buffer = ArrayPool<byte>.Shared.Rent(lengthToCopy);
            var target = new ExcelUtf8RangeRowSource(owner, buffer, lengthToCopy) {
                _firstColumn = _firstColumn + firstOrdinal,
                _fieldCount = fieldCount, _rowCount = count,
                _mappingRowsQualified = true,
                _mappingBufferFromSharedPool = true,
#if NET8_0_OR_GREATER
                _canonicalTextHasNonSimpleBytes = _canonicalTextHasNonSimpleBytes,
                _validatedCanonicalTextLength = maximum <= _validatedCanonicalTextLength ? lengthToCopy : 0,
#endif
            };
            try {
                if (lengthToCopy > 0) Buffer.BlockCopy(_buffer!, minimum, buffer, 0, lengthToCopy);
                target._rowIndexes = ExcelArrayPool.RentInts(Math.Max(1, count));
                int cells = checked(Math.Max(1, count) * fieldCount);
                target._valueStarts = ExcelArrayPool.RentInts(cells);
                target._valueLengths = ExcelArrayPool.RentInts(cells);
                target._cellKinds = ExcelArrayPool.RentBytes(cells);
                physical = firstPhysical;
                for (int row = 0; row < count; row++) {
                    ct.ThrowIfCancellationRequested();
                    target._rowIndexes[row] = row + 1;
                    int to = row * fieldCount;
                    target.InitializeMetadataRow(to);
                    if (physical >= _rowCount || _rowIndexes![physical] != firstRow + row) continue;
                    int from = physical * _fieldCount + firstOrdinal;
                    for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                        int originalLength = _valueLengths![from + ordinal];
                        target._valueLengths[to + ordinal] = originalLength;
                        target._valueStarts[to + ordinal] = originalLength < 0 ? 0 : _valueStarts![from + ordinal] - minimum;
                        target._cellKinds[to + ordinal] = _cellKinds![from + ordinal];
                    }
                    physical++;
                }
                copy = target;
                return true;
            } catch { target.Dispose(); throw; }
        }
    }
}
