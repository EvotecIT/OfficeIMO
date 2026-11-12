#if NET8_0_OR_GREATER
using System.Buffers;
using System.Text;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Recognizes only the five predefined XML entities. Numeric references
            /// and all other forms retain the complete XML/HTML-decoding fallback.
            /// This recognition never substitutes for structural XML qualification.
            /// </summary>
            private static bool TryReadPredefinedXmlEntity(ReadOnlySpan<byte> text, out char value, out int consumed) {
                value = default;
                consumed = 0;
                if (text.Length < 4) return false;
                switch (text[1]) {
                    case (byte)'a':
                        if (text.StartsWith("&amp;"u8)) { value = '&'; consumed = 5; }
                        else if (text.StartsWith("&apos;"u8)) { value = '\''; consumed = 6; }
                        break;
                    case (byte)'l':
                        if (text.StartsWith("&lt;"u8)) { value = '<'; consumed = 4; }
                        break;
                    case (byte)'g':
                        if (text.StartsWith("&gt;"u8)) { value = '>'; consumed = 4; }
                        break;
                    case (byte)'q':
                        if (text.StartsWith("&quot;"u8)) { value = '"'; consumed = 6; }
                        break;
                }
                return consumed != 0;
            }

            /// <summary>
            /// Materializes ASCII text with predefined entities directly into its final string in one
            /// pass. Literal CR, Unicode and numeric references return false and keep their
            /// established normalization and decoding. The source remains immutable throughout,
            /// after full worksheet validation has completed.
            /// </summary>
            private bool TryDecodePredefinedXmlEntities(int start, int length, out string text) {
                text = string.Empty;
                ReadOnlySpan<byte> input = _buffer!.AsSpan(start, length);
                // ASCII decodes to one character per byte, and each entity shortens the text.
                char[]? rented = null;
                Span<char> destination = length <= 256
                    ? stackalloc char[256]
                    : (rented = ArrayPool<char>.Shared.Rent(length));
                try {
                    int written = 0;
                    int cursor = 0;
                    while (true) {
                        int relative = input.Slice(cursor).IndexOfAny((byte)'&', (byte)'\r');
                        int end = relative < 0 ? length : cursor + relative;
                        if (Ascii.ToUtf16(input.Slice(cursor, end - cursor), destination.Slice(written), out int segment)
                            != OperationStatus.Done) {
                            return false;
                        }
                        written += segment;
                        if (relative < 0) break;
                        if (input[end] != (byte)'&'
                            || !TryReadPredefinedXmlEntity(input.Slice(end), out char value, out int consumed)) {
                            return false;
                        }
                        destination[written++] = value;
                        cursor = end + consumed;
                    }
                    text = new string(destination.Slice(0, written));
                    return true;
                } finally {
                    if (rented != null) ArrayPool<char>.Shared.Return(rented);
                }
            }
        }
    }
}
#endif
