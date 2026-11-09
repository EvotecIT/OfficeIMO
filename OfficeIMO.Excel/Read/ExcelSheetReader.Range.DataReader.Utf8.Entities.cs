#if NET8_0_OR_GREATER
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
                if (text.StartsWith("&amp;"u8)) { value = '&'; consumed = 5; }
                else if (text.StartsWith("&lt;"u8)) { value = '<'; consumed = 4; }
                else if (text.StartsWith("&gt;"u8)) { value = '>'; consumed = 4; }
                else if (text.StartsWith("&quot;"u8)) { value = '"'; consumed = 6; }
                else if (text.StartsWith("&apos;"u8)) { value = '\''; consumed = 6; }
                return consumed != 0;
            }

            /// <summary>
            /// Materializes common ASCII entity text directly into its final string.
            /// Literal CR, Unicode and numeric references keep their established
            /// normalization and decoding. The source remains immutable throughout
            /// both passes, after full worksheet validation has completed.
            /// </summary>
            private bool TryDecodePredefinedXmlEntities(int start, int length, out string text) {
                text = string.Empty;
                ReadOnlySpan<byte> input = _buffer!.AsSpan(start, length);
                int first = input.IndexOf((byte)'&');
                if (first < 0 || input.IndexOf((byte)'\r') >= 0 || !Ascii.IsValid(input)) return false;
                int decodedLength = length;
                int cursor = first;
                while (cursor < length) {
                    if (!TryReadPredefinedXmlEntity(input.Slice(cursor), out _, out int consumed)) return false;
                    decodedLength -= consumed - 1;
                    cursor += consumed;
                    int relative = input.Slice(cursor).IndexOf((byte)'&');
                    if (relative < 0) break;
                    cursor += relative;
                }
                text = string.Create(decodedLength, (Bytes: _buffer!, Start: start, Length: length),
                    static (destination, source) => {
                        ReadOnlySpan<byte> remaining = source.Bytes.AsSpan(source.Start, source.Length);
                        while (true) {
                            int entity = remaining.IndexOf((byte)'&');
                            if (entity < 0) {
                                _ = Ascii.ToUtf16(remaining, destination, out _);
                                return;
                            }
                            _ = Ascii.ToUtf16(remaining.Slice(0, entity), destination, out int written);
                            destination = destination.Slice(written);
                            if (!TryReadPredefinedXmlEntity(remaining.Slice(entity), out char value, out int consumed)) {
                                throw new InvalidOperationException("Qualified XML entity changed during decoding.");
                            }
                            destination[0] = value;
                            destination = destination.Slice(1);
                            remaining = remaining.Slice(entity + consumed);
                        }
                    });
                return true;
            }
        }
    }
}
#endif
