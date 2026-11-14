using System.Globalization;
using System.Text;
using XCellValues = DocumentFormat.OpenXml.Spreadsheet.CellValues;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Reads the rows and cells of a spooled streaming <c>sheetData</c> straight from its UTF-8 bytes. It yields the
        /// parts <see cref="AutoFitCellParts.Read"/> yields for the same markup, decoding entities and normalizing line
        /// breaks as an XML reader does, and reports anything else (shared strings, rich text, comments, unknown children)
        /// as unsupported so the caller measures the sheet through the XML reader instead.
        /// </summary>
        internal sealed class AutoFitSpoolScanner {
            internal enum Token { Row, Cell, End, Unsupported }

            private readonly Stream _stream;
            private byte[] _buffer = new byte[64 * 1024];
            private int _start;
            private int _end;
            private bool _eof;
            private char[] _chars = new char[256];

            internal AutoFitSpoolScanner(Stream stream) => _stream = stream;

            /// <summary>Moves to the next row or cell. A row token carries its number when the row states one.</summary>
            internal Token Next(out int rowNumber, out AutoFitCellParts parts) {
                rowNumber = 0;
                parts = default;
                while (true) {
                    if (!TrySkipTo((byte)'<')) return Token.End;
                    // Offsets inside an element stay valid while it reads, so unread bytes move to the front only between
                    // elements and only once half the buffer is consumed.
                    if (_start > _buffer.Length / 2) Compact();
                    int open = _start;
                    if (!TryFind((byte)'>', open + 1, out int tagEnd)) return Token.Unsupported;
                    int nameEnd = NameEnd(open + 1, tagEnd);
                    int nameLength = nameEnd - open - 1;
                    bool selfClosing = _buffer[tagEnd - 1] == (byte)'/';
                    if (IsName(open + 1, nameLength, "c")) {
                        return ReadCell(nameEnd, tagEnd, selfClosing, out parts);
                    }

                    if (IsName(open + 1, nameLength, "row")) {
                        rowNumber = ReadRowNumber(nameEnd, tagEnd);
                        _start = tagEnd + 1;
                        return Token.Row;
                    }

                    if (IsName(open + 1, nameLength, "/row") || IsName(open + 1, nameLength, "sheetData") || IsName(open + 1, nameLength, "/sheetData")) {
                        _start = tagEnd + 1;
                        continue;
                    }

                    return Token.Unsupported;
                }
            }

            private Token ReadCell(int attributesStart, int tagEnd, bool selfClosing, out AutoFitCellParts parts) {
                parts = default;
                uint styleIndex = 0U;
                XCellValues? dataType = null;
                int attributesEnd = selfClosing ? tagEnd - 1 : tagEnd;
                int position = attributesStart;
                while (TryReadAttribute(ref position, attributesEnd, out int nameStart, out int nameLength, out int valueStart, out int valueLength)) {
                    if (IsName(nameStart, nameLength, "s")) {
                        if (!TryParseUInt(valueStart, valueLength, out styleIndex)) return Token.Unsupported;
                    } else if (IsName(nameStart, nameLength, "t")) {
                        if (IsName(valueStart, valueLength, "n")) dataType = XCellValues.Number;
                        else if (IsName(valueStart, valueLength, "b")) dataType = XCellValues.Boolean;
                        else if (IsName(valueStart, valueLength, "e")) dataType = XCellValues.Error;
                        else if (IsName(valueStart, valueLength, "str")) dataType = XCellValues.String;
                        else if (IsName(valueStart, valueLength, "inlineStr")) dataType = XCellValues.InlineString;
                        else if (IsName(valueStart, valueLength, "d")) dataType = XCellValues.Date;
                        else return Token.Unsupported;
                    }
                }

                if (position < 0) return Token.Unsupported;
                if (selfClosing) {
                    _start = tagEnd + 1;
                    parts = new AutoFitCellParts(styleIndex, dataType, string.Empty, null);
                    return Token.Cell;
                }

                string value = string.Empty;
                string? inlineText = null;
                position = tagEnd + 1;
                while (true) {
                    if (!TryFind((byte)'<', position, out int open) || !TryFind((byte)'>', open, out int close)) return Token.Unsupported;
                    int childNameEnd = NameEnd(open + 1, close);
                    int childNameLength = childNameEnd - open - 1;
                    bool childSelfClosing = _buffer[close - 1] == (byte)'/';
                    if (IsName(open + 1, childNameLength, "/c")) {
                        _start = close + 1;
                        parts = new AutoFitCellParts(styleIndex, dataType, value, inlineText);
                        return Token.Cell;
                    }

                    if (IsName(open + 1, childNameLength, "v")) {
                        if (childSelfClosing) {
                            value = string.Empty;
                            position = close + 1;
                        } else if (!TryReadText(close + 1, "</v>", out value, out position)) {
                            return Token.Unsupported;
                        }
                    } else if (IsName(open + 1, childNameLength, "f")) {
                        if (childSelfClosing) position = close + 1;
                        else if (!TryReadText(close + 1, "</f>", out _, out position)) return Token.Unsupported;
                    } else if (IsName(open + 1, childNameLength, "is") && inlineText == null) {
                        if (childSelfClosing) {
                            inlineText = string.Empty;
                            position = close + 1;
                        } else if (!TryReadInlineText(close + 1, out inlineText, out position)) {
                            return Token.Unsupported;
                        }
                    } else {
                        return Token.Unsupported;
                    }
                }
            }

            /// <summary>Reads an inline string with one plain <c>t</c> element, or none.</summary>
            private bool TryReadInlineText(int position, out string? text, out int next) {
                text = string.Empty;
                next = position;
                if (!TryFind((byte)'<', position, out int open) || !TryFind((byte)'>', open, out int close)) return false;
                int nameEnd = NameEnd(open + 1, close);
                int nameLength = nameEnd - open - 1;
                if (IsName(open + 1, nameLength, "/is")) {
                    next = close + 1;
                    return true;
                }

                if (!IsName(open + 1, nameLength, "t")) return false;
                if (_buffer[close - 1] == (byte)'/') {
                    position = close + 1;
                } else if (!TryReadText(close + 1, "</t>", out text, out position)) {
                    return false;
                }

                if (!TryFind((byte)'<', position, out open) || !TryFind((byte)'>', open, out close)) return false;
                nameEnd = NameEnd(open + 1, close);
                if (!IsName(open + 1, nameEnd - open - 1, "/is")) return false;
                next = close + 1;
                return true;
            }

            /// <summary>Decodes element text up to its end tag; markup inside the text is not supported.</summary>
            private bool TryReadText(int position, string endTag, out string text, out int next) {
                text = string.Empty;
                next = position;
                if (!TryFind((byte)'<', position, out int open)) return false;
                if (!EnsureAvailable(open + endTag.Length - 1) || !IsName(open, endTag.Length, endTag)) return false;
                if (!TryDecode(position, open - position, out text)) return false;
                next = open + endTag.Length;
                return true;
            }

            private bool TryDecode(int start, int length, out string text) {
                text = string.Empty;
                if (length == 0) return true;
                int maxChars = Encoding.UTF8.GetMaxCharCount(length);
                if (_chars.Length < maxChars) _chars = new char[Math.Max(maxChars, _chars.Length * 2)];
                int count = Encoding.UTF8.GetChars(_buffer, start, length, _chars, 0);
                int write = 0;
                for (int read = 0; read < count; read++) {
                    char current = _chars[read];
                    if (current == '\r') {
                        // An XML reader turns literal CR LF and lone CR into LF.
                        if (read + 1 < count && _chars[read + 1] == '\n') read++;
                        _chars[write++] = '\n';
                    } else if (current == '&') {
                        int semicolon = Array.IndexOf(_chars, ';', read + 1, Math.Min(12, count - read - 1));
                        if (semicolon < 0 || !TryDecodeEntity(read + 1, semicolon - read - 1, ref write)) return false;
                        read = semicolon;
                    } else {
                        _chars[write++] = current;
                    }
                }

                text = new string(_chars, 0, write);
                return true;
            }

            private bool TryDecodeEntity(int start, int length, ref int write) {
                string name = new string(_chars, start, length);
                switch (name) {
                    case "amp": _chars[write++] = '&'; return true;
                    case "lt": _chars[write++] = '<'; return true;
                    case "gt": _chars[write++] = '>'; return true;
                    case "quot": _chars[write++] = '"'; return true;
                    case "apos": _chars[write++] = '\''; return true;
                }

                if (length < 2 || name[0] != '#') return false;
                bool hex = name[1] == 'x';
                if (!int.TryParse(hex ? name.Substring(2) : name.Substring(1), hex ? NumberStyles.AllowHexSpecifier : NumberStyles.None,
                        CultureInfo.InvariantCulture, out int codePoint) || codePoint <= 0 || codePoint > 0x10FFFF) {
                    return false;
                }

                string decoded = char.ConvertFromUtf32(codePoint);
                // Decoded text never outgrows its entity, so it fits where the entity was read from.
                foreach (char item in decoded) _chars[write++] = item;
                return true;
            }

            private int ReadRowNumber(int position, int tagEnd) {
                int end = _buffer[tagEnd - 1] == (byte)'/' ? tagEnd - 1 : tagEnd;
                while (TryReadAttribute(ref position, end, out int nameStart, out int nameLength, out int valueStart, out int valueLength)) {
                    if (IsName(nameStart, nameLength, "r") && TryParseUInt(valueStart, valueLength, out uint row) && row <= int.MaxValue) {
                        return (int)row;
                    }
                }

                return 0;
            }

            /// <summary>Reads one double-quoted attribute; position becomes negative for markup it does not expect.</summary>
            private bool TryReadAttribute(ref int position, int end, out int nameStart, out int nameLength, out int valueStart, out int valueLength) {
                nameStart = nameLength = valueStart = valueLength = 0;
                while (position < end && IsWhiteSpace(_buffer[position])) position++;
                if (position >= end) return false;
                nameStart = position;
                while (position < end && _buffer[position] != (byte)'=' && !IsWhiteSpace(_buffer[position])) position++;
                nameLength = position - nameStart;
                while (position < end && IsWhiteSpace(_buffer[position])) position++;
                if (position >= end || _buffer[position] != (byte)'=') {
                    position = -1;
                    return false;
                }

                position++;
                while (position < end && IsWhiteSpace(_buffer[position])) position++;
                if (position >= end || (_buffer[position] != (byte)'"' && _buffer[position] != (byte)'\'')) {
                    position = -1;
                    return false;
                }

                byte quote = _buffer[position++];
                valueStart = position;
                while (position < end && _buffer[position] != quote) position++;
                if (position >= end) {
                    position = -1;
                    return false;
                }

                valueLength = position - valueStart;
                position++;
                return true;
            }

            private bool TryParseUInt(int start, int length, out uint value) {
                value = 0U;
                if (length == 0 || length > 10) return false;
                ulong result = 0UL;
                for (int i = start; i < start + length; i++) {
                    int digit = _buffer[i] - (byte)'0';
                    if (digit < 0 || digit > 9) return false;
                    result = (result * 10UL) + (ulong)digit;
                }

                if (result > uint.MaxValue) return false;
                value = (uint)result;
                return true;
            }

            private bool IsName(int start, int length, string name) {
                if (length != name.Length) return false;
                for (int i = 0; i < length; i++) {
                    if (_buffer[start + i] != (byte)name[i]) return false;
                }

                return true;
            }

            /// <summary>The end of a tag name starting at <paramref name="start"/>, which may begin with the slash of an end tag.</summary>
            private int NameEnd(int start, int tagEnd) {
                int position = start < tagEnd && _buffer[start] == (byte)'/' ? start + 1 : start;
                while (position < tagEnd && _buffer[position] != (byte)'/' && !IsWhiteSpace(_buffer[position])) position++;
                return position;
            }

            private static bool IsWhiteSpace(byte value) => value == (byte)' ' || value == (byte)'\t' || value == (byte)'\r' || value == (byte)'\n';

            /// <summary>Skips to the next byte equal to <paramref name="value"/>, setting the start there.</summary>
            private bool TrySkipTo(byte value) {
                while (true) {
                    int index = Array.IndexOf(_buffer, value, _start, _end - _start);
                    if (index >= 0) {
                        _start = index;
                        return true;
                    }

                    _start = _end = 0;
                    if (!Fill()) return false;
                }
            }

            /// <summary>Finds <paramref name="value"/> at or after an offset from the start, reading more input as needed.</summary>
            private bool TryFind(byte value, int offset, out int index) {
                int from = offset;
                while (true) {
                    if (from < _end) {
                        index = Array.IndexOf(_buffer, value, from, _end - from);
                        if (index >= 0) return true;
                        from = _end;
                    }

                    if (!Fill()) {
                        index = -1;
                        return false;
                    }
                }
            }

            private bool EnsureAvailable(int index) {
                while (index >= _end) {
                    if (!Fill()) return false;
                }

                return true;
            }

            /// <summary>Moves the unread bytes to the front of the buffer.</summary>
            private void Compact() {
                if (_start == 0) return;
                int remaining = _end - _start;
                if (remaining > 0) Buffer.BlockCopy(_buffer, _start, _buffer, 0, remaining);
                _start = 0;
                _end = remaining;
            }

            private bool Fill() {
                if (_eof) return false;
                if (_end == _buffer.Length) Array.Resize(ref _buffer, _buffer.Length * 2);
                int read = _stream.Read(_buffer, _end, _buffer.Length - _end);
                if (read <= 0) {
                    _eof = true;
                    return false;
                }

                _end += read;
                return true;
            }
        }
    }
}
