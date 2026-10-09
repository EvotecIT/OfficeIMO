#if NET8_0_OR_GREATER
using System.Buffers;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        public sealed partial class ExcelTabularRowWriter {
            private static readonly UTF8Encoding StrictCellUtf8 = new(false, true);
            private static readonly SearchValues<byte> Utf8TextSpecialBytes = SearchValues.Create(
                new byte[] { 0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, (byte)'&', (byte)'<', (byte)'>' });

            /// <summary>Writes a text cell from a complete UTF-8 value on .NET 8 and later.</summary>
            /// <param name="value">Well-formed UTF-8 text containing at most 32,767 UTF-16 characters.</param>
            /// <returns>The current row writer.</returns>
            /// <remarks>
            /// Inline text is copied without creating a string. XML markup and carriage returns
            /// are escaped; unsupported XML control characters are removed as in Write(string).
            /// Shared-string output materializes the text for its string-interning table.
            /// The input is consumed during this call and is not retained.
            /// </remarks>
            /// <exception cref="ArgumentException">The text is invalid UTF-8, contains an invalid XML scalar, or exceeds Excel's text limit.</exception>
            public ExcelTabularRowWriter WriteUtf8(ReadOnlySpan<byte> value) {
                ValidateUtf8Text(value);
                if (_sharedStrings != null || _writer is not PooledUtf8TextWriter writer) {
                    return Write(StrictCellUtf8.GetString(value));
                }

                BeginCell();
                if (value.IsEmpty) {
                    _writer.Write(" t=\"inlineStr\"><is><t/></is></c>");
                    return this;
                }

                _writer.Write(" t=\"inlineStr\"><is><t");
                Rune.DecodeFromUtf8(value, out Rune first, out _);
                Rune.DecodeLastFromUtf8(value, out Rune last, out _);
                if (Rune.IsWhiteSpace(first) || Rune.IsWhiteSpace(last)) {
                    _writer.Write(" xml:space=\"preserve\"");
                }
                _writer.Write('>');
                WriteEscapedUtf8Text(writer, value);
                _writer.Write("</t></is></c>");
                return this;
            }

            private static void ValidateUtf8Text(ReadOnlySpan<byte> value) {
                int characters;
                try {
                    characters = StrictCellUtf8.GetCharCount(value);
                } catch (DecoderFallbackException exception) {
                    throw new ArgumentException("Text must contain well-formed UTF-8.", nameof(value), exception);
                }
                if (characters > CoerceValueHelper.SharedStringCharacterLimit) {
                    throw new ArgumentException("String exceeds Excel's limit of 32,767 characters.", nameof(value));
                }
                if (value.IndexOf("\uFFFE"u8) >= 0 || value.IndexOf("\uFFFF"u8) >= 0) {
                    throw new ArgumentException("Text contains a scalar that XML does not permit.", nameof(value));
                }
            }

            private static void WriteEscapedUtf8Text(PooledUtf8TextWriter writer, ReadOnlySpan<byte> value) {
                while (!value.IsEmpty) {
                    int special = value.IndexOfAny(Utf8TextSpecialBytes);
                    if (special < 0) {
                        writer.WriteUtf8(value);
                        return;
                    }
                    if (special > 0) writer.WriteUtf8(value.Slice(0, special));
                    switch (value[special]) {
                        case (byte)'&': writer.Write("&amp;"); break;
                        case (byte)'<': writer.Write("&lt;"); break;
                        case (byte)'>': writer.Write("&gt;"); break;
                        case (byte)'\r': writer.Write("&#xD;"); break;
                    }
                    value = value.Slice(special + 1);
                }
            }
        }
    }
}
#endif
