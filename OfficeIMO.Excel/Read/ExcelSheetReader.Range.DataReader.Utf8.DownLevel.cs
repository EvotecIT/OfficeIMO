#nullable enable

#if !NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private sealed partial class ExcelUtf8RangeRowSource {
            private const ulong RepeatedOnes = 0x0101010101010101UL;
            private const ulong LowSevenBits = 0x7F7F7F7F7F7F7F7FUL;
            private const ulong HighBits = 0x8080808080808080UL;
            private const int DenseRowMissesBeforeReview = 32;

            private int _explicitRowHits;
            private int _explicitRowMisses;
            private bool _explicitRowsDisabled;

            /// <summary>
            /// Indexes a row of the cell shape Excel writes, <c>&lt;c r="A1"[ s="n"][ t="s"]&gt;</c>
            /// with a plain <c>&lt;v&gt;</c> value, in one pass over local state. Any other shape,
            /// and any value the shared checks would reject, resets the row and hands it to the
            /// shared indexer, which then produces the same result or error it always did.
            /// </summary>
            private bool TryIndexDenseRowDownLevel(ref int position, int rowIndex, int rowOffset, CancellationToken ct) {
                if (_explicitRowsDisabled || _worksheetPrefixLength != 0 || _isParallelIndexHelper) {
                    return TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
                }

                int rowStart = position;
                if (TryIndexExplicitCellRow(ref position, rowIndex, rowOffset, ct)) {
                    _explicitRowHits++;
                    return true;
                }

                // Rows of other shapes would pay for two passes; stop trying when they are common.
                if (++_explicitRowMisses >= DenseRowMissesBeforeReview && _explicitRowMisses * 4 > _explicitRowHits) {
                    _explicitRowsDisabled = true;
                }
                position = rowStart;
                InitializeMetadataRow(rowOffset);
                return TryIndexCanonicalDenseRow(ref position, rowIndex, rowOffset, ct);
            }

            private bool TryIndexExplicitCellRow(ref int position, int rowIndex, int rowOffset, CancellationToken ct) {
                byte[] buffer = _buffer!;
                int end = _length;
                int p = position;
                int previousColumn = 0;
                int cellsUntilCancellationCheck = 0;
                while (true) {
                    if (p > end - 6) {
                        return false;
                    }
                    if (buffer[p] == (byte)'<' && buffer[p + 1] == (byte)'/') {
                        if (buffer[p + 2] != (byte)'r' || buffer[p + 3] != (byte)'o'
                            || buffer[p + 4] != (byte)'w' || buffer[p + 5] != (byte)'>') {
                            return false;
                        }
                        UpdateUsedBounds(rowIndex, previousColumn);
                        position = p + 6;
                        return true;
                    }
                    if (cellsUntilCancellationCheck-- == 0) {
                        ct.ThrowIfCancellationRequested();
                        cellsUntilCancellationCheck = 256;
                    }

                    // <c r="LETTERS DIGITS"
                    if (p > end - 7
                        || buffer[p] != (byte)'<' || buffer[p + 1] != (byte)'c' || buffer[p + 2] != (byte)' '
                        || buffer[p + 3] != (byte)'r' || buffer[p + 4] != (byte)'=' || buffer[p + 5] != (byte)'"') {
                        return false;
                    }
                    p += 6;
                    int column = 0;
                    int letters = 0;
                    while (p < end) {
                        byte current = buffer[p];
                        int letter = current >= (byte)'A' && current <= (byte)'Z'
                            ? current - (byte)'A' + 1
                            : current >= (byte)'a' && current <= (byte)'z'
                                ? current - (byte)'a' + 1
                                : 0;
                        if (letter == 0) {
                            break;
                        }
                        if (column > (int.MaxValue - letter) / 26) {
                            return false;
                        }
                        column = (column * 26) + letter;
                        letters++;
                        p++;
                    }
                    if (letters == 0 || p >= end || (uint)(buffer[p] - (byte)'0') > 9) {
                        return false;
                    }
                    do {
                        p++;
                    } while (p < end && (uint)(buffer[p] - (byte)'0') <= 9);
                    if (p >= end - 1 || buffer[p] != (byte)'"') {
                        return false;
                    }
                    p++;

                    // [ s="DIGITS"][ t="s"] then > or />
                    int styleIndex = -1;
                    if (p <= end - 5 && buffer[p] == (byte)' ' && buffer[p + 1] == (byte)'s'
                        && buffer[p + 2] == (byte)'=' && buffer[p + 3] == (byte)'"') {
                        p += 4;
                        int styleStart = p;
                        while (p < end && (uint)(buffer[p] - (byte)'0') <= 9) {
                            p++;
                        }
                        if (p >= end - 1 || buffer[p] != (byte)'"'
                            || !TryParseNonNegativeInt(buffer, styleStart, p - styleStart, out styleIndex)) {
                            return false;
                        }
                        p++;
                    }
                    Utf8CellKind kind = Utf8CellKind.Number;
                    if (p <= end - 7 && buffer[p] == (byte)' ' && buffer[p + 1] == (byte)'t' && buffer[p + 2] == (byte)'='
                        && buffer[p + 3] == (byte)'"' && buffer[p + 4] == (byte)'s' && buffer[p + 5] == (byte)'"') {
                        kind = Utf8CellKind.SharedString;
                        p += 6;
                    }
                    bool isEmpty = false;
                    if (buffer[p] == (byte)'/') {
                        if (p >= end - 1 || buffer[p + 1] != (byte)'>') {
                            return false;
                        }
                        isEmpty = true;
                        p++;
                    } else if (buffer[p] != (byte)'>') {
                        return false;
                    }
                    p++;

                    if (column <= previousColumn) {
                        return false;
                    }
                    if (previousColumn == 0
                        && (_minimumCellColumn == int.MaxValue || column < _minimumCellColumn)) {
                        _minimumCellColumn = column;
                    }
                    previousColumn = column;
                    int ordinal = column - _firstColumn;
                    if ((uint)ordinal >= (uint)_fieldCount) {
                        return false;
                    }
                    int cellIndex = rowOffset + ordinal;

                    // </c> or <v>TEXT</v></c>
                    int valueStart = -1;
                    int valueLength = -1;
                    if (!isEmpty) {
                        if (p > end - 4 || buffer[p] != (byte)'<') {
                            return false;
                        }
                        if (buffer[p + 1] == (byte)'/' && buffer[p + 2] == (byte)'c' && buffer[p + 3] == (byte)'>') {
                            p += 4;
                        } else if (buffer[p + 1] == (byte)'v' && buffer[p + 2] == (byte)'>') {
                            int contentStart = p + 3;
                            int valueEnd = contentStart;
                            while (valueEnd < end && buffer[valueEnd] != (byte)'<') {
                                valueEnd++;
                            }
                            if (valueEnd > end - 8
                                || buffer[valueEnd + 1] != (byte)'/' || buffer[valueEnd + 2] != (byte)'v'
                                || buffer[valueEnd + 3] != (byte)'>' || buffer[valueEnd + 4] != (byte)'<'
                                || buffer[valueEnd + 5] != (byte)'/' || buffer[valueEnd + 6] != (byte)'c'
                                || buffer[valueEnd + 7] != (byte)'>') {
                                return false;
                            }
                            valueStart = contentStart;
                            valueLength = valueEnd - contentStart;
                            p = valueEnd + 8;
                        } else {
                            return false;
                        }
                    }

                    // The shared indexer raises the errors for missing styles and shared strings.
                    if (styleIndex >= 0 && (uint)styleIndex >= (uint)_owner.Styles.CellFormatCount) {
                        return false;
                    }
                    if (kind == Utf8CellKind.SharedString) {
                        if (valueLength <= 0 || valueLength > 9) {
                            return false;
                        }
                        int sharedStringIndex = 0;
                        for (int index = valueStart; index < valueStart + valueLength; index++) {
                            uint digit = (uint)(buffer[index] - (byte)'0');
                            if (digit > 9) {
                                return false;
                            }
                            sharedStringIndex = (sharedStringIndex * 10) + (int)digit;
                        }
                        if ((uint)sharedStringIndex >= (uint)_owner._sst.Count) {
                            return false;
                        }
                        _valueStarts![cellIndex] = sharedStringIndex;
                        _valueLengths![cellIndex] = SharedStringIndexValueLength;
                    } else if (valueStart >= 0) {
                        _valueStarts![cellIndex] = valueStart;
                        _valueLengths![cellIndex] = valueLength;
                    }
                    _cellKinds![cellIndex] = EncodeCellKind(kind, styleIndex);
                }
            }

            /// <summary>
            /// Proves the worksheet text is UTF-8 without XML-forbidden characters, entities or a
            /// CDATA terminator. Eight bytes are examined at a time; only bytes that could break the
            /// proof (non-ASCII, control characters, '&amp;' and ']') are looked at one by one.
            /// </summary>
            private bool IsCanonicalTextValidDownLevel(CancellationToken ct) {
                byte[] buffer = _buffer!;
                int length = _length;
                bool utf8Validated = false;
                int index = 0;
                int nextCancellationCheck = 0;
                while (true) {
                    if (index >= nextCancellationCheck) {
                        ct.ThrowIfCancellationRequested();
                        nextCancellationCheck = index + 0x10000;
                    }
                    index = IndexOfTextCandidate(buffer, index, length);
                    if (index < 0) {
                        return true;
                    }

                    byte value = buffer[index];
                    if (value >= 0x80) {
                        if (!utf8Validated) {
                            // Everything before the first non-ASCII byte is ASCII, so this is a
                            // sequence boundary and one strict decode proves the remainder.
                            try {
                                _ = StrictUtf8.GetCharCount(buffer, index, length - index);
                            } catch (DecoderFallbackException) {
                                return false;
                            }
                            utf8Validated = true;
                        }
                        if (value == 0xef
                            && index + 2 < length
                            && buffer[index + 1] == 0xbf
                            && buffer[index + 2] is 0xbe or 0xbf) {
                            return false;
                        }
                    } else if (value == (byte)'&'
                               || value < 0x20 && value is not (byte)'\t' and not (byte)'\r' and not (byte)'\n'
                               || value == (byte)']'
                                   && index + 2 < length
                                   && buffer[index + 1] == (byte)']'
                                   && buffer[index + 2] == (byte)'>') {
                        return false;
                    }
                    index++;
                }
            }

            /// <summary>
            /// Compares a short tag literal in place. Literals of four to eight bytes are compared as
            /// two overlapping 32-bit words; bounds are checked for truncated documents.
            /// </summary>
            private bool MatchesUtf8(int start, ReadOnlySpan<byte> value) {
                int length = value.Length;
                if (start < 0 || start > _length - length) {
                    return false;
                }
                ref byte actual = ref _buffer![start];
                ref byte expected = ref MemoryMarshal.GetReference(value);
                if (length >= sizeof(uint) && length <= sizeof(ulong)) {
                    int last = length - sizeof(uint);
                    return Unsafe.ReadUnaligned<uint>(ref actual) == Unsafe.ReadUnaligned<uint>(ref expected)
                        && Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref actual, last))
                            == Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref expected, last));
                }
                for (int index = 0; index < length; index++) {
                    if (Unsafe.Add(ref actual, index) != Unsafe.Add(ref expected, index)) {
                        return false;
                    }
                }
                return true;
            }

            /// <summary>Finds the next byte that is non-ASCII, a control character, '&amp;' or ']'.</summary>
            private static int IndexOfTextCandidate(byte[] buffer, int start, int end) {
                int index = start;
                if (end - index >= sizeof(ulong)) {
                    ref byte data = ref buffer[0];
                    int lastWord = end - sizeof(ulong);
                    for (; index <= lastWord; index += sizeof(ulong)) {
                        ulong word = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref data, index));
                        ulong flags = (word & HighBits)
                            | ControlBytes(word)
                            | MatchingBytes(word ^ ((byte)'&' * RepeatedOnes))
                            | MatchingBytes(word ^ ((byte)']' * RepeatedOnes));
                        if (flags != 0) {
                            return index + LowestFlaggedByte(flags);
                        }
                    }
                }

                for (; index < end; index++) {
                    byte value = buffer[index];
                    if (value >= 0x80 || value < 0x20 || value == (byte)'&' || value == (byte)']') {
                        return index;
                    }
                }
                return -1;
            }

            /// <summary>Sets the high bit of every byte below 0x20 (ASCII bytes only; callers flag the rest).</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static ulong ControlBytes(ulong word) =>
                ~((word & LowSevenBits) + (0x60 * RepeatedOnes)) & ~word & HighBits;

            /// <summary>Sets the high bit of every zero byte, exactly (no borrow between bytes).</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static ulong MatchingBytes(ulong word) => ~(((word & LowSevenBits) + LowSevenBits) | word | LowSevenBits);

            /// <summary>Index of the lowest byte whose high bit is set (little-endian word).</summary>
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            private static int LowestFlaggedByte(ulong flags) {
                ulong lowest = flags & (~flags + 1);
                // Each flag is a byte's top bit; (bit >> 7) * pattern puts the byte number in the top byte.
                return (int)(((lowest >> 7) * 0x0001020304050607UL) >> 56);
            }
        }
    }
}
#endif
