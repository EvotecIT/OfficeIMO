#nullable enable

#if NET8_0_OR_GREATER

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader : IDataReaderUtf8TextSource {
        private bool[] _borrowableText = null!;
        private int[] _utf8TextCacheKeys = null!;
        private ExcelUtf8TextCache? _utf8TextCache;

        partial void InitializeUtf8TextState() {
            _borrowableText = _options.InferSchema || _options.CellValueConverter != null
                ? Array.Empty<bool>()
                : new bool[FieldCount];
            _utf8TextCacheKeys = _borrowableText.Length == 0
                ? Array.Empty<int>()
                : new int[FieldCount];
        }

        partial void ResetUtf8TextState() {
            Array.Clear(_borrowableText);
            Array.Clear(_utf8TextCacheKeys);
        }

        partial void TrackUtf8TextCell(int ordinal, int recordType) {
            if (_borrowableText.Length != 0) {
                _borrowableText[ordinal] = recordType is BrtCellSt or BrtCellRString;
                // Shared-string indexes are nonnegative. Inline keys occupy the other namespace;
                // canonical string identity also protects a key reused by a later inline row.
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

            // Canonical XLSB strings are UTF-16. Encoding a cache miss allocates an immutable array;
            // slot replacement never changes the bytes already borrowed by another field.
            _utf8TextCache ??= new ExcelUtf8TextCache();
            text = _utf8TextCache.Get(_utf8TextCacheKeys[ordinal], _strings[ordinal]!).AsSpan();
            return true;
        }
    }
}
#endif
