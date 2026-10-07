using System.Text;

namespace OfficeIMO.Excel {
    internal sealed partial class SharedStringCache {
        /// <summary>
        /// Reads the small, plain ASCII SpreadsheetML shape without constructing an XML reader.
        /// Any markup, escaping, Unicode, or different attribute layout uses the full XML path.
        /// </summary>
        private bool TryLoadSimpleAsciiItems(OpenXmlPooledPartStream stream, out List<string> items) {
            items = null!;
            byte[] bytes = stream.BorrowBuffer(out int length);
            int position = 0;
            if (!Consume(bytes, length, ref position, "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")) {
                return false;
            }
            SkipWhitespace(bytes, length, ref position);
            if (!Consume(bytes, length, ref position,
                    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")) {
                return false;
            }
            if (!ReadUnsignedDecimal(bytes, length, ref position, out _)
                || !Consume(bytes, length, ref position, "\" uniqueCount=\"")
                || !ReadUnsignedDecimal(bytes, length, ref position, out int declaredUniqueCount)
                || !Consume(bytes, length, ref position, "\">")) {
                return false;
            }

            // Reject common full-XML cases before allocating any item strings.
            // A late Unicode value or entity otherwise discards the ASCII prefix
            // and allocates it again when the XML reader reparses the whole part.
#if NET8_0_OR_GREATER
            ReadOnlySpan<byte> content = bytes.AsSpan(position, length - position);
            if (!Ascii.IsValid(content) || content.IndexOf((byte)'&') >= 0) return false;
#else
            for (int index = position; index < length; index++) {
                if (bytes[index] >= 0x80 || bytes[index] == (byte)'&') return false;
            }
#endif

            var parsed = new List<string>(Math.Min(declaredUniqueCount,
                Math.Min(_maxSharedStringItems, length / 7)));
            long totalCharacters = 0;
            while (!Consume(bytes, length, ref position, "</sst>")) {
                _cancellationToken.ThrowIfCancellationRequested();
                bool preservesWhitespace = false;
                if (!Consume(bytes, length, ref position, "<si><t>")) {
                    if (!Consume(bytes, length, ref position, "<si><t xml:space=\"preserve\">")) {
                        return false;
                    }
                    preservesWhitespace = true;
                }

                int textStart = position;
                while (position < length && bytes[position] != (byte)'<') {
                    byte value = bytes[position];
                    if (value < 0x20 || value > 0x7E || value == (byte)'&'
                        || value == (byte)']' && position + 2 < length
                            && bytes[position + 1] == (byte)']'
                            && bytes[position + 2] == (byte)'>') {
                        return false;
                    }
                    position++;
                }

                int textLength = position - textStart;
                if (!preservesWhitespace && textLength > 0) {
                    bool onlyWhitespace = true;
                    for (int index = textStart; index < position; index++) {
                        if (bytes[index] != (byte)' ') {
                            onlyWhitespace = false;
                            break;
                        }
                    }
                    if (onlyWhitespace) return false;
                }
                if (!Consume(bytes, length, ref position, "</t></si>")) {
                    return false;
                }

                EnsureCanAddSharedString(parsed);
                EnsureItemCharacterBudget(0, textLength, _maxSharedStringItemCharacters);
                string valueText = Encoding.ASCII.GetString(bytes, textStart, textLength);
                ValidateSharedStringText(valueText, ref totalCharacters);
                parsed.Add(valueText);
            }

            SkipWhitespace(bytes, length, ref position);
            if (position != length) {
                return false;
            }

            _cancellationToken.ThrowIfCancellationRequested();
            items = parsed;
            return true;
        }

        private static bool Consume(byte[] bytes, int length, ref int position, string value) {
            if (value.Length > length - position) return false;
            for (int index = 0; index < value.Length; index++) {
                if (bytes[position + index] != (byte)value[index]) return false;
            }
            position += value.Length;
            return true;
        }

        private static bool ReadUnsignedDecimal(byte[] bytes, int length, ref int position, out int value) {
            value = 0;
            int start = position;
            while (position < length && bytes[position] >= (byte)'0' && bytes[position] <= (byte)'9') {
                int digit = bytes[position++] - (byte)'0';
                if (value > (int.MaxValue - digit) / 10) return false;
                value = value * 10 + digit;
            }
            return position > start;
        }

        private static void SkipWhitespace(byte[] bytes, int length, ref int position) {
            while (position < length && bytes[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') {
                position++;
            }
        }
    }
}
