#nullable enable

#if NET8_0_OR_GREATER
using System.Globalization;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSchemaInferenceDataReader : IDataReaderUtf8TextSource {
        private bool _hasUtf8TextRow;

        partial void SetUtf8TextCursor(bool hasCurrentRow) => _hasUtf8TextRow = hasCurrentRow;

        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
            if (_closed) {
                throw new InvalidOperationException("The reader is closed.");
            }
            if ((uint)ordinal >= (uint)FieldCount) {
                throw new IndexOutOfRangeException(ordinal.ToString(CultureInfo.InvariantCulture));
            }
            if (!_hasUtf8TextRow) {
                throw new InvalidOperationException("The reader is not positioned on a row.");
            }

            // The inner reader may already point beyond a replayed sample row.
            // Schema replay exposes materialized objects rather than borrowed bytes.
            text = default;
            return false;
        }
    }
}
#endif
