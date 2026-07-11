#nullable enable

namespace OfficeIMO.Excel {
    public sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
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
                    int relative = _buffer!.AsSpan(position, limit - position).IndexOf((byte)'<');
                    if (relative < 0) {
                        position = limit;
                        return false;
                    }

                    int start = position + relative;
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
                while (position < limit && IsAsciiWhitespace(_buffer![position])) position++;
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
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position] == (byte)'/') return true;
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
                    while (position < tag.End && IsAsciiWhitespace(_buffer![position])) position++;
                    if (position >= tag.End || _buffer![position] == (byte)'/') break;
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
                    } else if (AsciiEquals(attributeStart, attributeEnd - attributeStart, "s")
                        && TryParseNonNegativeInt(_buffer!, valueStart, valueLength, out int parsedStyle)) {
                        styleIndex = parsedStyle;
                    }

                    position++;
                }

                if (hasReference) {
                    columnIndex = ParseColumnIndex(_buffer!, referenceStart, referenceLength);
                    if (columnIndex <= 0) return false;
                    nextColumn = columnIndex + 1;
                } else {
                    columnIndex = nextColumn++;
                }

                return !hasType || TryParseCellKind(typeStart, typeLength, out kind);
            }

            private bool TryParseCellKind(int start, int length, out Utf8CellKind kind) {
                kind = Utf8CellKind.Number;
                if (AsciiEquals(start, length, "n")) return true;
                if (AsciiEquals(start, length, "s")) { kind = Utf8CellKind.SharedString; return true; }
                if (AsciiEquals(start, length, "str")) { kind = Utf8CellKind.String; return true; }
                if (AsciiEquals(start, length, "b")) { kind = Utf8CellKind.Boolean; return true; }
                if (AsciiEquals(start, length, "d")) { kind = Utf8CellKind.Date; return true; }
                if (AsciiEquals(start, length, "e")) { kind = Utf8CellKind.Error; return true; }
                return false;
            }
        }
    }
}