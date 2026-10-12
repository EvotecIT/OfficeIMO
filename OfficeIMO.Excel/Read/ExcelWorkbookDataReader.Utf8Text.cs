#nullable enable

#if NET8_0_OR_GREATER

namespace OfficeIMO.Excel {
    public sealed partial class ExcelWorkbookDataReader : IDataReaderUtf8TextSource {
        /// <summary>Attempts to borrow the current field's text from reader-owned UTF-8 storage.</summary>
        /// <remarks>
        /// Plain UTF-8 worksheet text, normalized shared-string text, and XLSB string cells can be borrowed.
        /// Canonical shared strings and XLSB strings may allocate UTF-8 storage on a cache miss;
        /// the reader caches small values within fixed bounds.
        /// Fields requiring worksheet XML decoding, formula projection, schema replay, or cell conversion
        /// return false. Consume the span before advancing to another row or result,
        /// or closing or disposing the reader.
        /// Copy the bytes when they must outlive the current row. Use GetString when this returns false.
        /// </remarks>
        /// <param name="ordinal">The zero-based field ordinal.</param>
        /// <param name="text">The borrowed UTF-8 text on success; otherwise an empty span.</param>
        /// <returns>True when the field is text that can be borrowed from reader-owned storage.</returns>
        /// <exception cref="IndexOutOfRangeException">The ordinal is outside the reader's fields.</exception>
        /// <exception cref="InvalidOperationException">The reader is closed or is not positioned on a row.</exception>
        public bool TryGetUtf8Text(int ordinal, out ReadOnlySpan<byte> text) {
            ThrowIfClosed();
            if (_current is IDataReaderUtf8TextSource source) {
                return source.TryGetUtf8Text(ordinal, out text);
            }

            if ((uint)ordinal >= (uint)FieldCount) {
                throw new IndexOutOfRangeException(ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            text = default;
            return false;
        }
    }
}
#endif
