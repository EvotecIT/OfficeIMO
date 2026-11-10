using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel;

internal sealed partial class ExcelSheetReader {
    private sealed partial class ExcelUtf8RangeRowSource {
        private MappingStorageLease? _mappingStorage;
        private int _mappingMetadataOffset;

        /// <summary>Shares completed indexes for dense full-width batches, retaining the original pool ownership.</summary>
        private bool TryShareMappingRows(ExcelSheetReader owner, int firstRow, int count,
            int firstPhysical, int firstOrdinal, int fieldCount, CancellationToken ct,
            out ExcelUtf8RangeRowSource? rows) {
            rows = null;
            if (count <= 0 || firstOrdinal != 0 || fieldCount != _fieldCount || !_ownsBuffer
                || firstPhysical > _rowCount - count
                || _rowIndexes![firstPhysical] != firstRow
                || _rowIndexes[firstPhysical + count - 1] != firstRow + count - 1) return false;

            // Strictly increasing validated row indexes make equal endpoint/count spans contiguous.
            int metadataOffset = checked(firstPhysical * fieldCount);
            var target = new ExcelUtf8RangeRowSource(owner, _buffer!, _length) {
                _ownsBuffer = false,
                _firstColumn = _firstColumn,
                _fieldCount = fieldCount,
                _rowCount = count,
                _mappingRowsQualified = true,
                _mappingMetadataOffset = metadataOffset,
#if NET8_0_OR_GREATER
                _canonicalTextHasNonSimpleBytes = _canonicalTextHasNonSimpleBytes,
                _validatedCanonicalTextLength = _validatedCanonicalTextLength,
#endif
            };
            try {
                target._rowIndexes = ArrayPool<int>.Shared.Rent(count);
                for (int row = 0; row < count; row++) {
                    ct.ThrowIfCancellationRequested();
                    target._rowIndexes[row] = row + 1;
                }

                var storage = _mappingStorage ??= new MappingStorageLease(
                    _buffer!, _valueStarts!, _valueLengths!, _cellKinds!, _mappingBufferFromSharedPool);
                storage.AddReference();
                target._mappingStorage = storage;
                target._valueStarts = _valueStarts;
                target._valueLengths = _valueLengths;
                target._cellKinds = _cellKinds;
                rows = target;
                return true;
            } catch { target.Dispose(); throw; }
        }

        /// <summary>Detaches this reader's shared references before the last lease returns their arrays.</summary>
        private void ReleaseSharedMappingStorage() {
            var storage = Interlocked.Exchange(ref _mappingStorage, null);
            if (storage == null) return;
            _buffer = null;
            _valueStarts = null;
            _valueLengths = null;
            _cellKinds = null;
            storage.Release();
        }

        /// <summary>Keeps immutable worksheet payloads and indexes alive through parent and worker disposal.</summary>
        private sealed class MappingStorageLease {
            private readonly byte[] _buffer;
            private readonly int[] _starts;
            private readonly int[] _lengths;
            private readonly byte[] _kinds;
            private readonly bool _sharedBufferPool;
            private int _references = 1;

            internal MappingStorageLease(byte[] buffer, int[] starts, int[] lengths, byte[] kinds,
                bool sharedBufferPool) {
                _buffer = buffer;
                _starts = starts;
                _lengths = lengths;
                _kinds = kinds;
                _sharedBufferPool = sharedBufferPool;
            }

            internal void AddReference() => Interlocked.Increment(ref _references);

            internal void Release() {
                if (Interlocked.Decrement(ref _references) != 0) return;
                if (_sharedBufferPool) ArrayPool<byte>.Shared.Return(_buffer, clearArray: true);
                else OpenXmlPartBufferPool.Return(_buffer);
                ArrayPool<int>.Shared.Return(_starts);
                ArrayPool<int>.Shared.Return(_lengths);
                ArrayPool<byte>.Shared.Return(_kinds);
            }
        }
    }
}
