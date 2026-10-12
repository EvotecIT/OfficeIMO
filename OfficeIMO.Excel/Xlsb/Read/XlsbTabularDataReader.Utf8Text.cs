#nullable enable

#if NET8_0_OR_GREATER

using System.Buffers;
using System.Text;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader : IDataReaderUtf8TextSource {
        private const int MaximumRetainedUtf8TextBytes = ExcelUtf8TextCache.SlotCount
            * ExcelUtf8TextCache.MaximumCachedItemBytes;
        private bool[] _borrowableText = null!;
        private int[] _utf8TextCacheKeys = null!;
        private ExcelUtf8TextCache? _utf8TextCache;
        private Utf8FieldBuffer[]? _utf8FieldBuffers;
        private bool _useSharedStringUtf8Cache;
        private int _maximumRetainedUtf8FieldBytes;
        private int _retainedUtf8FieldBytes;

        partial void InitializeUtf8TextState() {
            _borrowableText = _options.InferSchema || _options.CellValueConverter != null
                ? Array.Empty<bool>()
                : new bool[FieldCount];
            _utf8TextCacheKeys = _borrowableText.Length == 0
                ? Array.Empty<int>()
                : new int[FieldCount];
            _useSharedStringUtf8Cache = _sharedStrings.Count <= ExcelUtf8TextCache.SlotCount;
            // Reserve the immutable cache's maximum payload before retaining row buffers.
            // Small SST indexes occupy distinct slots; inline text never enters that cache.
            int sharedStringReservation = _useSharedStringUtf8Cache
                ? _sharedStrings.Count * ExcelUtf8TextCache.MaximumCachedItemBytes
                : 0;
            _maximumRetainedUtf8FieldBytes = MaximumRetainedUtf8TextBytes - sharedStringReservation;
        }

        partial void ResetUtf8TextState() {
            Array.Clear(_borrowableText);
            Array.Clear(_utf8TextCacheKeys);
            if (_utf8FieldBuffers != null) {
                for (int ordinal = 0; ordinal < _utf8FieldBuffers.Length; ordinal++) {
                    ref Utf8FieldBuffer buffer = ref _utf8FieldBuffers[ordinal];
                    if (!buffer.Retained) {
                        ReturnUtf8FieldBuffer(ref buffer);
                    } else {
                        buffer.Length = 0;
                        buffer.Ready = false;
                    }
                }
            }
        }

        partial void TrackUtf8TextCell(int ordinal, int recordType) {
            if (_borrowableText.Length != 0) {
                _borrowableText[ordinal] = recordType is BrtCellSt or BrtCellRString;
                // Shared-string indexes are nonnegative. Inline keys select row-owned storage.
                _utf8TextCacheKeys[ordinal] = ~ordinal;
            }
        }

        partial void TrackUtf8SharedStringCell(int ordinal, int sharedStringIndex) {
            if (_borrowableText.Length != 0) {
                _borrowableText[ordinal] = true;
                _utf8TextCacheKeys[ordinal] = sharedStringIndex;
            }
        }

        partial void ReleaseUtf8TextState() {
            ResetUtf8TextState();
            if (_utf8FieldBuffers != null) {
                for (int ordinal = 0; ordinal < _utf8FieldBuffers.Length; ordinal++) {
                    ReturnUtf8FieldBuffer(ref _utf8FieldBuffers[ordinal]);
                }
                _utf8FieldBuffers = null;
            }
            _utf8TextCache = null;
        }

        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
            ValidateReadableOrdinal(ordinal);
            // Caller options can change, but a reader created without borrow state stays disabled.
            if (_borrowableText.Length == 0 || _options.InferSchema || _options.CellValueConverter != null
                || _kinds[ordinal] != XlsbTabularValueKind.Text || !_borrowableText[ordinal]) {
                text = default;
                return false;
            }

            int key = _utf8TextCacheKeys[ordinal];
            if (_useSharedStringUtf8Cache && key >= 0) {
                _utf8TextCache ??= new ExcelUtf8TextCache();
                text = _utf8TextCache.Get(key, _strings[ordinal]!).AsSpan();
                return true;
            }

            Utf8FieldBuffer[] buffers = _utf8FieldBuffers ??= new Utf8FieldBuffer[FieldCount];
            ref Utf8FieldBuffer buffer = ref buffers[ordinal];
            if (!buffer.Ready) {
                EncodeUtf8Field(ref buffer, _strings[ordinal]!);
            }
            text = buffer.Length == 0 ? ReadOnlySpan<byte>.Empty : buffer.Bytes!.AsSpan(0, buffer.Length);
            return true;
        }

        /// <summary>
        /// Encodes a field once per row. Each ordinal owns its storage, so another field or
        /// repeated accessor cannot overwrite a span borrowed earlier in the same row.
        /// </summary>
        private void EncodeUtf8Field(ref Utf8FieldBuffer buffer, string value) {
            int length = Encoding.UTF8.GetByteCount(value);
            if (length == 0) {
                buffer.Length = 0;
                buffer.Ready = true;
                return;
            }

            byte[]? bytes = buffer.Bytes;
            if (bytes == null || bytes.Length < length) {
                // This field has not been borrowed in the current row. Its prior-row bytes
                // may be returned even when other ordinals already have live spans.
                ReturnUtf8FieldBuffer(ref buffer);
                bytes = ArrayPool<byte>.Shared.Rent(length);
                buffer.Bytes = bytes;
                buffer.Retained = bytes.Length <= ExcelUtf8TextCache.MaximumCachedItemBytes
                    && bytes.Length <= _maximumRetainedUtf8FieldBytes - _retainedUtf8FieldBytes;
                if (buffer.Retained) {
                    _retainedUtf8FieldBytes += bytes.Length;
                }
            }

            buffer.Length = Encoding.UTF8.GetBytes(value.AsSpan(), bytes.AsSpan());
            buffer.Ready = true;
        }

        /// <summary>Returns owned bytes only after their current-row borrow has ended.</summary>
        private void ReturnUtf8FieldBuffer(ref Utf8FieldBuffer buffer) {
            byte[]? bytes = buffer.Bytes;
            if (bytes != null) {
                if (buffer.Retained) {
                    _retainedUtf8FieldBytes -= bytes.Length;
                }
                buffer = default;
                // Clear the whole rental, including a longer value left by an earlier row.
                ArrayPool<byte>.Shared.Return(bytes, clearArray: true);
            } else {
                buffer = default;
            }
        }

        private struct Utf8FieldBuffer {
            internal byte[]? Bytes;
            internal int Length;
            internal bool Ready;
            internal bool Retained;
        }
    }
}
#endif
