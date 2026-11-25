using System.Text;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private bool TryGetDeclaredRange(
                CancellationToken ct,
                out int firstRow,
                out int firstColumn,
                out int lastRow,
                out int lastColumn) {
                firstRow = 0;
                firstColumn = 0;
                lastRow = 0;
                lastColumn = 0;
                if (!HasSupportedUtf8Encoding()) {
                    return false;
                }

                int position = 0;
                while (TryReadNextTag(ref position, _length, out Utf8Tag tag)) {
                    ct.ThrowIfCancellationRequested();
                    if (tag.IsEnd) {
                        continue;
                    }

                    if (LocalNameEquals(tag, "sheetData")) {
                        return false;
                    }

                    if (!LocalNameEquals(tag, "dimension")
                        || !TryGetAttribute(tag, "ref", out bool found, out int valueStart, out int valueLength)
                        || !found) {
                        continue;
                    }

                    string reference = Encoding.UTF8.GetString(_buffer!, valueStart, valueLength);
                    return ExcelSheetReader.TryNormalizeWorksheetDimensionReference(reference, out string normalized)
                        && ExcelA1.TryParseRange(normalized, out firstRow, out firstColumn, out lastRow, out lastColumn);
                }

                return false;
            }

            private bool HasSupportedUtf8Encoding() {
                int probeLength = Math.Min(_length, 256);
                for (int i = 0; i < probeLength; i++) {
                    if (_buffer![i] == 0) {
                        return false;
                    }
                }

                int encoding = IndexOfAsciiIgnoreCase(0, probeLength, "encoding");
                if (encoding < 0) {
                    return true;
                }

                int position = encoding + 8;
                while (position < probeLength && IsAsciiWhitespace(_buffer![position])) position++;
                if (position >= probeLength || _buffer![position++] != (byte)'=') return false;
                while (position < probeLength && IsAsciiWhitespace(_buffer![position])) position++;
                if (position >= probeLength || (_buffer![position] != (byte)'\"' && _buffer[position] != (byte)'\'')) return false;
                byte quote = _buffer[position++];
                int start = position;
                while (position < probeLength && _buffer[position] != quote) position++;
                return position < probeLength
                    && (AsciiEqualsIgnoreCase(start, position - start, "utf-8") || AsciiEqualsIgnoreCase(start, position - start, "utf8"));
            }

            private bool TryReadNextTag(ref int position, int limit, out Utf8Tag tag) {
                tag = default;
                while (position < limit) {
                    int start;
                    if (_buffer![position] == (byte)'<') {
                        start = position;
                    } else {
                        int relative = _buffer.AsSpan(position, limit - position).IndexOf((byte)'<');
                        if (relative < 0) {
                            position = limit;
                            return false;
                        }
                        start = position + relative;
                    }
                    if (start + 1 >= limit) {
                        _parseFailed = true;
                        position = limit;
                        return false;
                    }

                    if (_buffer![start + 1] == (byte)'?') {
                        int end = IndexOfSequence(start + 2, limit, (byte)'?', (byte)'>');
                        if (end < 0) {
                            _parseFailed = true;
                            return false;
                        }

                        position = end + 2;
                        continue;
                    }

                    if (_buffer[start + 1] == (byte)'!') {
                        if (start + 3 < limit && _buffer[start + 2] == (byte)'-' && _buffer[start + 3] == (byte)'-') {
                            int end = IndexOfSequence(start + 4, limit, (byte)'-', (byte)'-', (byte)'>');
                            if (end < 0) {
                                _parseFailed = true;
                                return false;
                            }

                            position = end + 3;
                            continue;
                        }

                        _parseFailed = true;
                        return false;
                    }

                    if (!TryParseTag(start, limit, out tag)) {
                        _parseFailed = true;
                        return false;
                    }

                    position = tag.End + 1;
                    return true;
                }

                return false;
            }

            private bool TryParseTag(int start, int limit, out Utf8Tag tag) {
                tag = default;
                int position = start + 1;
                bool isEnd = position < limit && _buffer![position] == (byte)'/';
                if (isEnd) position++;
                if (position >= limit || IsAsciiWhitespace(_buffer![position])) return false;
                int nameStart = position;
                while (position < limit && !IsTagNameTerminator(_buffer![position])) position++;
                if (position == nameStart) return false;
                int nameEnd = position;

                byte quote = 0;
                while (position < limit) {
                    byte current = _buffer![position];
                    if (quote != 0) {
                        if (current == quote) quote = 0;
                    } else if (current == (byte)'\"' || current == (byte)'\'') {
                        quote = current;
                    } else if (current == (byte)'>') {
                        int beforeEnd = position - 1;
                        while (beforeEnd >= nameEnd && IsAsciiWhitespace(_buffer[beforeEnd])) beforeEnd--;
                        bool isEmpty = !isEnd && beforeEnd >= nameEnd && _buffer[beforeEnd] == (byte)'/';
                        if (isEnd) {
                            for (int index = nameEnd; index < position; index++) {
                                if (!IsAsciiWhitespace(_buffer[index])) return false;
                            }
                        } else if (isEmpty && beforeEnd != position - 1) {
                            return false;
                        }
                        int localNameStart = nameStart;
                        for (int i = nameStart; i < nameEnd; i++) {
                            if (_buffer[i] == (byte)':') localNameStart = i + 1;
                        }

                        tag = new Utf8Tag(start, position, nameStart, nameEnd, localNameStart, isEnd, isEmpty);
                        return true;
                    }

                    position++;
                }

                return false;
            }

            private bool TryGetAttribute(Utf8Tag tag, string name, out bool found, out int valueStart, out int valueLength) {
                found = false;
                valueStart = 0;
                valueLength = 0;
                int position = tag.NameEnd;
                while (position < tag.End) {
                    int whitespaceStart = position;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position] == (byte)'/') return true;
                    if (position == whitespaceStart) return false;
                    int attributeStart = position;
                    while (position < tag.End && !IsAttributeNameTerminator(_buffer![position])) position++;
                    int attributeEnd = position;
                    if (attributeEnd == attributeStart) return false;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position++] != (byte)'=') return false;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || (_buffer![position] != (byte)'\"' && _buffer[position] != (byte)'\'')) return false;
                    byte quote = _buffer[position++];
                    int start = position;
                    while (position < tag.End && _buffer[position] != quote) position++;
                    if (position >= tag.End) return false;

                    if (AsciiEquals(attributeStart, attributeEnd - attributeStart, name)) {
                        found = true;
                        valueStart = start;
                        valueLength = position - start;
                    }

                    position++;
                }

                return true;
            }

            private bool TryGetCellAttributes(
                Utf8Tag tag,
                ref int nextColumn,
                out int columnIndex,
                out Utf8CellKind kind,
                out int styleIndex) {
                columnIndex = 0;
                kind = Utf8CellKind.Number;
                styleIndex = -1;
                int referenceStart = 0;
                int referenceLength = 0;
                int typeStart = 0;
                int typeLength = 0;
                bool hasReference = false;
                bool hasType = false;
                int position = tag.NameEnd;
                while (position < tag.End) {
                    int whitespaceStart = position;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position] == (byte)'/') break;
                    if (position == whitespaceStart) return false;
                    int attributeStart = position;
                    while (position < tag.End && !IsAttributeNameTerminator(_buffer![position])) position++;
                    int attributeEnd = position;
                    if (attributeEnd == attributeStart) return false;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position++] != (byte)'=') return false;
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || (_buffer![position] != (byte)'\"' && _buffer[position] != (byte)'\'')) return false;
                    byte quote = _buffer[position++];
                    int valueStart = position;
                    while (position < tag.End && _buffer[position] != quote) position++;
                    if (position >= tag.End) return false;
                    int valueLength = position - valueStart;

                    if (AsciiEquals(attributeStart, attributeEnd - attributeStart, "r")) {
                        hasReference = true;
                        referenceStart = valueStart;
                        referenceLength = valueLength;
                    } else if (AsciiEquals(attributeStart, attributeEnd - attributeStart, "t")) {
                        hasType = true;
                        typeStart = valueStart;
                        typeLength = valueLength;
                    } else if (AsciiEquals(attributeStart, attributeEnd - attributeStart, "s")) {
                        if (!TryParseNonNegativeInt(_buffer!, valueStart, valueLength, out int parsedStyle)) {
                            if (valueLength != 0) return false;
                            // Excel repairs an empty style to the default style; keep the cell unstyled.
                            ReportEmptyCellStyle(hasReference, referenceStart, referenceLength);
                            parsedStyle = -1;
                        }
                        styleIndex = parsedStyle;
                    } else if (AsciiEquals(attributeStart, attributeEnd - attributeStart, "vm")) {
                        // Rich error metadata requires the shared XML decoder.
                        return false;
                    }

                    position++;
                }

                if (hasReference) {
                    columnIndex = ParseColumnIndex(_buffer!, referenceStart, referenceLength);
                    if (columnIndex <= 0
                        && !TryResolveUnparsedCellColumn(referenceStart, referenceLength, nextColumn, out columnIndex)) {
                        return false;
                    }
                    nextColumn = columnIndex + 1;
                } else {
                    columnIndex = nextColumn++;
                }

                return !hasType || TryParseCellKind(typeStart, typeLength, out kind);
            }

            /// <summary>
            /// Resolves a reference the byte parser rejected with the XmlReader paths' rule: the attribute text is
            /// decoded, then <see cref="GetXmlCellColumnIndex(ReadOnlySpan{char}, ref int)"/> trims and parses it and
            /// places an unparseable value in the next column. Single-pass reading cannot switch to the XmlReader after
            /// returning rows, so it needs the same result here. Text this decoder does not handle returns false.
            /// </summary>
            private bool TryResolveUnparsedCellColumn(int start, int length, int nextColumn, out int columnIndex) {
                columnIndex = 0;
                Span<char> text = length <= 64 ? stackalloc char[64] : new char[length];
                int written = 0;
                int position = 0;
                while (position < length) {
                    byte current = _buffer![start + position];
                    if (current >= 0x80) return false;
                    if (current != (byte)'&') {
                        text[written++] = (char)current;
                        position++;
                        continue;
                    }
                    if (!TryDecodeXmlCharacterReference(_buffer.AsSpan(start + position, length - position), out char decoded, out int consumed)) {
                        return false;
                    }
                    text[written++] = decoded;
                    position += consumed;
                }
                columnIndex = GetXmlCellColumnIndex(text.Slice(0, written), ref nextColumn);
                return true;
            }

            /// <summary>Decodes a numeric character reference or one of the five predefined XML entities.</summary>
            private static bool TryDecodeXmlCharacterReference(ReadOnlySpan<byte> text, out char value, out int consumed) {
                value = default;
                consumed = 0;
                int end = text.IndexOf((byte)';');
                if (end < 2) return false;
                ReadOnlySpan<byte> name = text.Slice(1, end - 1);
                consumed = end + 1;
                if (name[0] == (byte)'#') {
                    bool hex = name.Length > 1 && (name[1] == (byte)'x');
                    int digitsStart = hex ? 2 : 1;
                    if (digitsStart >= name.Length) return false;
                    int code = 0;
                    for (int i = digitsStart; i < name.Length; i++) {
                        byte digit = name[i];
                        int part = digit >= (byte)'0' && digit <= (byte)'9' ? digit - '0'
                            : hex && digit >= (byte)'a' && digit <= (byte)'f' ? digit - 'a' + 10
                            : hex && digit >= (byte)'A' && digit <= (byte)'F' ? digit - 'A' + 10
                            : -1;
                        if (part < 0) return false;
                        code = code * (hex ? 16 : 10) + part;
                        if (code > 0xFFFF) return false;
                    }
                    if (code == 0 || (code >= 0xD800 && code <= 0xDFFF)) return false;
                    value = (char)code;
                    return true;
                }
                if (AsciiNameEquals(name, "amp")) value = '&';
                else if (AsciiNameEquals(name, "lt")) value = '<';
                else if (AsciiNameEquals(name, "gt")) value = '>';
                else if (AsciiNameEquals(name, "quot")) value = '"';
                else if (AsciiNameEquals(name, "apos")) value = '\'';
                else return false;
                return true;
            }

            private static bool AsciiNameEquals(ReadOnlySpan<byte> name, string expected) {
                if (name.Length != expected.Length) return false;
                for (int i = 0; i < name.Length; i++) {
                    if (name[i] != (byte)expected[i]) return false;
                }
                return true;
            }

            private void ReportEmptyCellStyle(bool hasReference, int referenceStart, int referenceLength) {
                if (!_owner.TryBeginEmptyCellStyleRepairReport()) return;
                _owner.ReportEmptyCellStyleRepair(hasReference
                    ? Encoding.UTF8.GetString(_buffer!, referenceStart, referenceLength)
                    : "(unknown cell)");
            }

            private bool TryParseCellKind(int start, int length, out Utf8CellKind kind) {
                kind = Utf8CellKind.Number;
                if (AsciiEquals(start, length, "n")) return true;
                if (AsciiEquals(start, length, "s")) { kind = Utf8CellKind.SharedString; return true; }
                if (AsciiEquals(start, length, "inlineStr")) { kind = Utf8CellKind.InlineString; return true; }
                if (AsciiEquals(start, length, "str")) { kind = Utf8CellKind.String; return true; }
                if (AsciiEquals(start, length, "b")) { kind = Utf8CellKind.Boolean; return true; }
                if (AsciiEquals(start, length, "d")) { kind = Utf8CellKind.Date; return true; }
                if (AsciiEquals(start, length, "e")) { kind = Utf8CellKind.Error; return true; }
                return false;
            }
        }
    }
}
