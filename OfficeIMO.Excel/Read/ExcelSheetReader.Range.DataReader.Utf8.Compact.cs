#nullable enable

using System.Threading;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private bool TryCaptureRepeatedRowShape(Utf8Tag tag, int expectedRowIndex) {
                int cursor = tag.Start;
                if (!TryReadRepeatedRowPrefix(ref cursor, out int rowIndex)
                    || rowIndex != expectedRowIndex) {
                    return false;
                }

                int suffixLength = tag.End - cursor + 1;
                if (suffixLength <= 0) {
                    return false;
                }

                _repeatedRowSuffixStart = cursor;
                _repeatedRowSuffixLength = suffixLength;
                _repeatedRowIsEmpty = tag.IsEmpty;
                return true;
            }

            private bool TryReadRepeatedRowStartTag(
                ref int position,
                out Utf8Tag tag,
                out int rowIndex) {
                tag = default;
                rowIndex = 0;
                int start = position;
                int cursor = position;
                if (_repeatedRowSuffixStart < 0
                    || !TryReadRepeatedRowPrefix(ref cursor, out rowIndex)
                    || cursor > _length - _repeatedRowSuffixLength
                    || !_buffer!.AsSpan(cursor, _repeatedRowSuffixLength).SequenceEqual(
                        _buffer.AsSpan(_repeatedRowSuffixStart, _repeatedRowSuffixLength))) {
                    rowIndex = 0;
                    return false;
                }

                int end = cursor + _repeatedRowSuffixLength - 1;
                tag = new Utf8Tag(
                    start,
                    end,
                    start + 1,
                    start + 4,
                    start + 1,
                    isEnd: false,
                    _repeatedRowIsEmpty);
                position = end + 1;
                return true;
            }

            private bool TryReadRepeatedRowPrefix(ref int cursor, out int rowIndex) {
                rowIndex = 0;
                if (cursor > _length - 9
                    || _buffer![cursor] != (byte)'<'
                    || _buffer[cursor + 1] != (byte)'r'
                    || _buffer[cursor + 2] != (byte)'o'
                    || _buffer[cursor + 3] != (byte)'w'
                    || _buffer[cursor + 4] != (byte)' '
                    || _buffer[cursor + 5] != (byte)'r'
                    || _buffer[cursor + 6] != (byte)'='
                    || _buffer[cursor + 7] != (byte)'\"') {
                    return false;
                }

                cursor += 8;
                int referenceStart = cursor;
                while (cursor < _length && _buffer[cursor] is >= (byte)'0' and <= (byte)'9') {
                    cursor++;
                }
                int referenceLength = cursor - referenceStart;
                if (referenceLength == 0) {
                    return false;
                }

                rowIndex = ParsePositiveInt(_buffer, referenceStart, referenceLength);
                return rowIndex > 0;
            }

            private bool TryReadCompactCellStartTag(
                ref int position,
                ref int nextColumn,
                out Utf8Tag tag,
                out int columnIndex,
                out Utf8CellKind kind,
                out int styleIndex) {
                if (TryReadCanonicalCellStartTag(
                        ref position,
                        ref nextColumn,
                        out tag,
                        out columnIndex,
                        out kind,
                        out styleIndex)) {
                    return true;
                }

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
                    int whitespaceStart = cursor;
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
                        if (cursor >= _length || _buffer[cursor] != (byte)'>') {
                            return false;
                        }
                        isEmpty = true;
                        break;
                    }
                    if (cursor == whitespaceStart) {
                        return false;
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

            private bool TryIndexCanonicalDenseRow(
                ref int position,
                int rowIndex,
                int rowOffset,
                CancellationToken ct) {
                int nextColumn = 1;
                int previousColumn = 0;
                int cellsUntilCancellationCheck = 0;
                while (position < _length) {
                    if (position <= _length - 6
                        && _buffer![position + 1] == (byte)'/'
                        && MatchesAscii(position, "</row>")) {
                        position += 6;
                        UpdateUsedBounds(rowIndex, previousColumn);
                        return true;
                    }

                    if (cellsUntilCancellationCheck-- == 0) {
                        ct.ThrowIfCancellationRequested();
                        cellsUntilCancellationCheck = 256;
                    }

                    if (!TryReadCanonicalCellStartTag(
                            ref position,
                            ref nextColumn,
                            out Utf8Tag tag,
                            out int columnIndex,
                            out Utf8CellKind kind,
                            out int styleIndex)
                        || columnIndex <= previousColumn) {
                        return false;
                    }

                    if (previousColumn == 0
                        && (_minimumCellColumn == int.MaxValue || columnIndex < _minimumCellColumn)) {
                        _minimumCellColumn = columnIndex;
                    }
                    previousColumn = columnIndex;
                    int ordinal = columnIndex - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        return false;
                    }

                    int cellIndex = rowOffset + ordinal;
                    bool hasCachedValue = false;
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!tag.IsEmpty
                        && !TryIndexCompactValueCell(
                            ref position,
                            cellIndex,
                            out hasCachedValue,
                            out valueStart,
                            out valueLength)) {
                        return false;
                    }

                    ValidateIndexedCell(
                        rowIndex,
                        columnIndex,
                        kind,
                        styleIndex,
                        sharedFormulaFollower: false,
                        hasCachedValue,
                        valueStart,
                        valueLength);
                    _cellKinds![cellIndex] = EncodeCellKind(kind, styleIndex);
                }

                return false;
            }

            private bool TryReadCanonicalCellStartTag(
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

                int start = position;
                int cursor = start;
                if (cursor > _length - 7
                    || _buffer![cursor] != (byte)'<'
                    || _buffer[cursor + 1] != (byte)'c'
                    || _buffer[cursor + 2] != (byte)' '
                    || _buffer[cursor + 3] != (byte)'r'
                    || _buffer[cursor + 4] != (byte)'='
                    || _buffer[cursor + 5] != (byte)'\"') {
                    return false;
                }

                cursor += 6;
                int parsedColumn = 0;
                int letterCount = 0;
                while (cursor < _length) {
                    byte current = _buffer[cursor];
                    int letter = current >= (byte)'A' && current <= (byte)'Z'
                        ? current - (byte)'A' + 1
                        : current >= (byte)'a' && current <= (byte)'z'
                            ? current - (byte)'a' + 1
                            : 0;
                    if (letter == 0) {
                        break;
                    }
                    if (parsedColumn > (int.MaxValue - letter) / 26) {
                        return false;
                    }
                    parsedColumn = (parsedColumn * 26) + letter;
                    letterCount++;
                    cursor++;
                }

                if (letterCount == 0 || cursor >= _length || _buffer[cursor] is < (byte)'0' or > (byte)'9') {
                    return false;
                }
                do {
                    cursor++;
                } while (cursor < _length && _buffer[cursor] is >= (byte)'0' and <= (byte)'9');
                if (cursor >= _length || _buffer[cursor++] != (byte)'\"') {
                    return false;
                }
                if (cursor >= _length) {
                    return false;
                }

                bool isEmpty = false;
                if (_buffer[cursor] == (byte)'>') {
                    // Number cell without an explicit style or type.
                } else if (_buffer[cursor] == (byte)'/'
                           && cursor + 1 < _length
                           && _buffer[cursor + 1] == (byte)'>') {
                    isEmpty = true;
                    cursor++;
                } else if (cursor <= _length - 7
                           && _buffer[cursor] == (byte)' '
                           && _buffer[cursor + 1] == (byte)'t'
                           && _buffer[cursor + 2] == (byte)'='
                           && _buffer[cursor + 3] == (byte)'\"'
                           && _buffer[cursor + 4] == (byte)'s'
                           && _buffer[cursor + 5] == (byte)'\"'
                           && _buffer[cursor + 6] == (byte)'>') {
                    kind = Utf8CellKind.SharedString;
                    cursor += 6;
                } else if (cursor <= _length - 6
                           && _buffer[cursor] == (byte)' '
                           && _buffer[cursor + 1] == (byte)'s'
                           && _buffer[cursor + 2] == (byte)'='
                           && _buffer[cursor + 3] == (byte)'\"') {
                    cursor += 4;
                    int styleStart = cursor;
                    while (cursor < _length && _buffer[cursor] is >= (byte)'0' and <= (byte)'9') {
                        cursor++;
                    }
                    if (cursor == styleStart
                        || cursor + 1 >= _length
                        || _buffer[cursor++] != (byte)'\"'
                        || _buffer[cursor] != (byte)'>') {
                        return false;
                    }
                    if (!TryParseNonNegativeInt(_buffer, styleStart, cursor - styleStart - 1, out styleIndex)) {
                        return false;
                    }
                } else {
                    return false;
                }

                columnIndex = parsedColumn;
                tag = new Utf8Tag(start, cursor, start + 1, start + 2, start + 1, isEnd: false, isEmpty);
                nextColumn = parsedColumn + 1;
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
