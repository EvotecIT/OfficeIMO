#nullable enable

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            /// <summary>
            /// Recognizes ordinary unprefixed cell tags with implicit columns. Exact
            /// attribute syntax proves their XML shape; unusual attributes and explicit
            /// coordinates retain the existing parser and coordinate validation.
            /// </summary>
            private bool TryReadCanonicalImplicitCellStartTag(
                ref int position,
                ref int nextColumn,
                out Utf8Tag tag,
                out int columnIndex,
                out Utf8CellKind kind,
                out int styleIndex) {
                tag = default;
                columnIndex = 0;
                int start = position;
                if (!TryReadCanonicalImplicitCellAttributes(ref position, out kind, out styleIndex, out bool isEmpty)) {
                    return false;
                }

                columnIndex = nextColumn;
                tag = new Utf8Tag(start, position - 1, start + 1, start + 2, start + 1, isEnd: false, isEmpty);
                nextColumn = columnIndex + 1;
                return true;
            }

            // Dense implicit rows need the attributes and empty marker, but none of
            // the general tag coordinates or an explicit column reference.
            private bool TryReadCanonicalImplicitCellAttributes(
                ref int position,
                out Utf8CellKind kind,
                out int styleIndex,
                out bool isEmpty) {
                kind = Utf8CellKind.Number;
                styleIndex = -1;
                isEmpty = false;
                int start = position;
                if (start > _length - 3
                    || _buffer![start] != (byte)'<'
                    || _buffer[start + 1] != (byte)'c') {
                    return false;
                }

                int cursor = start + 2;
                bool sawType = false;
                bool sawStyle = false;
                while (cursor < _length && _buffer[cursor] == (byte)' ') {
                    if (cursor > _length - 5
                        || _buffer[cursor + 2] != (byte)'='
                        || _buffer[cursor + 3] != (byte)'"') {
                        return false;
                    }
                    byte attribute = _buffer[cursor + 1];
                    cursor += 4;
                    int valueStart = cursor;
                    if (attribute == (byte)'s' && !sawStyle) {
                        while (cursor < _length && _buffer[cursor] is >= (byte)'0' and <= (byte)'9') {
                            cursor++;
                        }
                        if (cursor == valueStart
                            || !TryParseNonNegativeInt(_buffer, valueStart, cursor - valueStart, out styleIndex)) {
                            return false;
                        }
                        sawStyle = true;
                    } else if (attribute == (byte)'t' && !sawType) {
                        int relativeEnd = _buffer.AsSpan(cursor, _length - cursor).IndexOf((byte)'"');
                        if (relativeEnd < 0 || !TryParseCellKind(valueStart, relativeEnd, out kind)) {
                            return false;
                        }
                        cursor += relativeEnd;
                        sawType = true;
                    } else {
                        return false;
                    }
                    if (cursor >= _length || _buffer[cursor++] != (byte)'"') {
                        return false;
                    }
                }

                isEmpty = cursor < _length && _buffer[cursor] == (byte)'/';
                if (isEmpty) cursor++;
                if (cursor >= _length || _buffer[cursor] != (byte)'>') {
                    return false;
                }

                position = cursor + 1;
                return true;
            }

            /// <summary>
            /// Indexes simple inline text while proving every enclosing end tag. Rich
            /// text, namespaces, text attributes and other constructs use the general
            /// parser. Full document byte validation still checks UTF-8, entities and
            /// XML characters before the indexed source is exposed.
            /// </summary>
            private bool TryIndexCanonicalInlineStringCell(
                ref int position,
                int cellIndex,
                out int valueStart,
                out int valueLength) {
                valueStart = valueLength = -1;
                int cursor = position;
                if (!MatchesUtf8(cursor, "<is><t>"u8)) {
                    return false;
                }

                int contentStart = cursor + 7;
                int relativeEnd = _buffer!.AsSpan(contentStart, _length - contentStart).IndexOf((byte)'<');
                if (relativeEnd < 0) {
                    return false;
                }
                int valueEnd = contentStart + relativeEnd;
                if (!MatchesUtf8(valueEnd, "</t></is></c>"u8)) {
                    return false;
                }

                valueStart = contentStart;
                valueLength = valueEnd - contentStart;
                if (cellIndex >= 0) {
                    _valueStarts![cellIndex] = valueStart;
                    _valueLengths![cellIndex] = valueLength;
                }
                position = valueEnd + 13;
                return true;
            }
        }
    }
}
