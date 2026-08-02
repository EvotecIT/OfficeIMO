#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private bool TryReadCompactCellStartTag(
                ref int position,
                ref int nextColumn,
                out Utf8Tag tag,
                out int columnIndex,
                out Utf8CellKind kind,
                out int styleIndex) {
                tag = default;
                columnIndex = 0;
                kind = Utf8CellKind.Number;
                styleIndex = -1;
                int cursor = position;
                if (cursor + 2 >= _length
                    || _buffer![cursor] != (byte)'<'
                    || _buffer[cursor + 1] != (byte)'c'
                    || !IsTagNameTerminator(_buffer[cursor + 2])) {
                    return false;
                }

                int start = cursor;
                cursor += 2;
                int referenceStart = 0;
                int referenceLength = 0;
                int typeStart = 0;
                int typeLength = 0;
                bool hasReference = false;
                bool hasType = false;
                bool hasStyle = false;
                bool isEmpty = false;
                while (cursor < _length) {
                    while (cursor < _length && IsAsciiWhitespace(_buffer[cursor])) {
                        cursor++;
                    }
                    if (cursor >= _length) {
                        return false;
                    }
                    if (_buffer[cursor] == (byte)'>') {
                        break;
                    }
                    if (_buffer[cursor] == (byte)'/') {
                        cursor++;
                        while (cursor < _length && IsAsciiWhitespace(_buffer[cursor])) {
                            cursor++;
                        }
                        if (cursor >= _length || _buffer[cursor] != (byte)'>') {
                            return false;
                        }
                        isEmpty = true;
                        break;
                    }

                    int attributeStart = cursor;
                    while (cursor < _length && !IsAttributeNameTerminator(_buffer[cursor])) {
                        cursor++;
                    }
                    int attributeLength = cursor - attributeStart;
                    if (attributeLength == 0) {
                        return false;
                    }
                    while (cursor < _length && IsAsciiWhitespace(_buffer[cursor])) {
                        cursor++;
                    }
                    if (cursor >= _length || _buffer[cursor++] != (byte)'=') {
                        return false;
                    }
                    while (cursor < _length && IsAsciiWhitespace(_buffer[cursor])) {
                        cursor++;
                    }
                    if (cursor >= _length || _buffer[cursor] is not ((byte)'\"') and not ((byte)'\'')) {
                        return false;
                    }
                    byte quote = _buffer[cursor++];
                    int valueStart = cursor;
                    while (cursor < _length && _buffer[cursor] != quote) {
                        if (_buffer[cursor] == (byte)'<') {
                            return false;
                        }
                        cursor++;
                    }
                    if (cursor >= _length) {
                        return false;
                    }
                    int valueLength = cursor - valueStart;
                    cursor++;

                    if (AsciiEquals(attributeStart, attributeLength, "r")) {
                        if (hasReference) {
                            return false;
                        }
                        hasReference = true;
                        referenceStart = valueStart;
                        referenceLength = valueLength;
                    } else if (AsciiEquals(attributeStart, attributeLength, "t")) {
                        if (hasType) {
                            return false;
                        }
                        hasType = true;
                        typeStart = valueStart;
                        typeLength = valueLength;
                    } else if (AsciiEquals(attributeStart, attributeLength, "s")) {
                        if (hasStyle) {
                            return false;
                        }
                        hasStyle = true;
                        if (!TryParseNonNegativeInt(_buffer, valueStart, valueLength, out styleIndex)) {
                            return false;
                        }
                    } else {
                        return false;
                    }
                }

                if (cursor >= _length || _buffer[cursor] != (byte)'>') {
                    return false;
                }
                if (hasReference) {
                    columnIndex = ParseColumnIndex(_buffer, referenceStart, referenceLength);
                    if (columnIndex <= 0) {
                        return false;
                    }
                } else {
                    columnIndex = nextColumn;
                }
                if (hasType && !TryParseCellKind(typeStart, typeLength, out kind)) {
                    return false;
                }

                tag = new Utf8Tag(start, cursor, start + 1, start + 2, start + 1, isEnd: false, isEmpty);
                nextColumn = columnIndex + 1;
                position = cursor + 1;
                return true;
            }

            private bool TryIndexCompactValueCell(
                ref int position,
                int cellIndex,
                out bool hasCachedValue,
                out int valueStart,
                out int valueLength) {
                hasCachedValue = false;
                valueStart = -1;
                valueLength = -1;
                int cursor = position;
                if (MatchesAscii(cursor, "</c>")) {
                    position = cursor + 4;
                    return true;
                }
                if (!MatchesAscii(cursor, "<v>")) {
                    return false;
                }

                int contentStart = cursor + 3;
                int relativeEnd = _buffer!.AsSpan(contentStart, _length - contentStart).IndexOf((byte)'<');
                if (relativeEnd < 0) {
                    return false;
                }
                int valueEnd = contentStart + relativeEnd;
                if (!MatchesAscii(valueEnd, "</v>") || !MatchesAscii(valueEnd + 4, "</c>")) {
                    return false;
                }

                valueStart = contentStart;
                valueLength = valueEnd - contentStart;
                hasCachedValue = true;
                if (cellIndex >= 0) {
                    _valueStarts![cellIndex] = valueStart;
                    _valueLengths![cellIndex] = valueLength;
                }
                position = valueEnd + 8;
                return true;
            }

            private bool MatchesAscii(int start, string value) =>
                start >= 0
                && start <= _length - value.Length
                && AsciiEquals(start, value.Length, value);
        }
    }
}
