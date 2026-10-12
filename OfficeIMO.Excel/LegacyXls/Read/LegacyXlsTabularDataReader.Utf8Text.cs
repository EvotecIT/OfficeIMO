#nullable enable

#if NET8_0_OR_GREATER

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader : IDataReaderUtf8TextSource {
        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
            ValidateReadableOrdinal(ordinal);
            text = default;
            return false;
        }
    }
}
#endif
