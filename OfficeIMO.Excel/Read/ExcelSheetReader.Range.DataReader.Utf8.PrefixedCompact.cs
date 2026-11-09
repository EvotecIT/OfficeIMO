namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            // Qualified compact tags prove exact body syntax. Character and
            // inherited namespace proof remain mandatory before publication.
            private bool MatchesPrefixedCompactTag(int start, ReadOnlySpan<byte> tag) {
                int leadingLength = tag[1] == (byte)'/' ? 2 : 1;
                int prefixLength = _worksheetPrefixLength;
                if (start < 0 || start > _length - tag.Length - prefixLength) return false;
                byte[] buffer = _buffer!;
                return buffer.AsSpan(start, leadingLength).SequenceEqual(tag.Slice(0, leadingLength))
                    && buffer.AsSpan(start + leadingLength, prefixLength).SequenceEqual(
                        buffer.AsSpan(_worksheetPrefixStart, prefixLength))
                    && buffer.AsSpan(start + leadingLength + prefixLength, tag.Length - leadingLength)
                        .SequenceEqual(tag.Slice(leadingLength));
            }

            private bool TryConsumeCompactRowEnd(ref int position) {
                int prefixLength = _worksheetPrefixLength;
                if (prefixLength == 0 ? !MatchesUtf8(position, "</row>"u8)
                    : !MatchesPrefixedCompactTag(position, "</row>"u8)) return false;
                position += 6 + prefixLength;
                return true;
            }

            private bool TryIndexPrefixedCompactValueCell(
                ref int position,
                int cellIndex,
                Utf8CellKind kind,
                out bool hasCachedValue,
                out int valueStart,
                out int valueLength) {
                hasCachedValue = false;
                valueStart = valueLength = -1;
                int cursor = position;
                int prefixLength = _worksheetPrefixLength;
                if (MatchesPrefixedCompactTag(cursor, "</c>"u8)) {
                    position = cursor + 4 + prefixLength;
                    return true;
                }

                bool inlineString = kind == Utf8CellKind.InlineString;
                int contentStart;
                if (inlineString) {
                    if (!MatchesPrefixedCompactTag(cursor, "<is>"u8)
                        || !MatchesPrefixedCompactTag(cursor + 4 + prefixLength, "<t>"u8)) return false;
                    contentStart = cursor + 7 + 2 * prefixLength;
                } else {
                    if (!MatchesPrefixedCompactTag(cursor, "<v>"u8)) return false;
                    contentStart = cursor + 3 + prefixLength;
                }

                int relativeEnd = _buffer!.AsSpan(contentStart, _length - contentStart).IndexOf((byte)'<');
                if (relativeEnd < 0) return false;
                int valueEnd = contentStart + relativeEnd;
                int cellEnd;
                if (inlineString) {
                    if (!MatchesPrefixedCompactTag(valueEnd, "</t>"u8)
                        || !MatchesPrefixedCompactTag(valueEnd + 4 + prefixLength, "</is>"u8)) return false;
                    cellEnd = valueEnd + 9 + 2 * prefixLength;
                } else {
                    if (!MatchesPrefixedCompactTag(valueEnd, "</v>"u8)) return false;
                    cellEnd = valueEnd + 4 + prefixLength;
                }
                if (!MatchesPrefixedCompactTag(cellEnd, "</c>"u8)) return false;

                valueStart = contentStart;
                valueLength = valueEnd - contentStart;
                hasCachedValue = !inlineString;
                if (cellIndex >= 0) {
                    _valueStarts![cellIndex] = valueStart;
                    _valueLengths![cellIndex] = valueLength;
                }
                position = cellEnd + 4 + prefixLength;
                return true;
            }
        }
    }
}
