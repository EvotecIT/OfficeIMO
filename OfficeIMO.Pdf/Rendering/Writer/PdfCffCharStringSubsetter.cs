namespace OfficeIMO.Pdf;

/// <summary>
/// Rewrites a CFF1 CharStrings INDEX while preserving glyph identifiers.
/// </summary>
/// <remarks>
/// PDF content streams use the original glyph identifiers. Keeping the INDEX count and order stable
/// avoids a second glyph-remapping layer: used programs remain byte-for-byte identical, glyph zero is
/// retained, and unused programs become the minimal Type 2 <c>endchar</c> program. Local and global
/// subroutines remain intact because retained programs may reference them.
/// </remarks>
internal static class PdfCffCharStringSubsetter {
    private const byte Type2EndChar = 14;

    internal static PdfCffCharStringSubset Create(byte[] cffData, IReadOnlyList<int> usedGlyphIds, int expectedGlyphCount) {
        Guard.NotNull(cffData, nameof(cffData));
        Guard.NotNull(usedGlyphIds, nameof(usedGlyphIds));

        try {
            if (cffData.Length < 4 || cffData[0] != 1) {
                return PdfCffCharStringSubset.Unsupported(cffData, "Only CFF1 data can be subset through the OpenType CFF path.");
            }

            int headerSize = cffData[2];
            if (headerSize < 4 || headerSize > cffData.Length) {
                return PdfCffCharStringSubset.Unsupported(cffData, "The CFF header size is invalid.");
            }

            CffIndex nameIndex = ReadIndex(cffData, headerSize);
            CffIndex topDictionaryIndex = ReadIndex(cffData, nameIndex.EndOffset);
            if (topDictionaryIndex.Count != 1) {
                return PdfCffCharStringSubset.Unsupported(cffData, "Only single-font CFF1 data can be subset.");
            }

            int charStringsOffset = FindTopDictionaryInteger(cffData, topDictionaryIndex.GetObject(0), 17);
            if (charStringsOffset <= 0 || charStringsOffset >= cffData.Length) {
                return PdfCffCharStringSubset.Unsupported(cffData, "The CFF Top DICT does not expose a valid CharStrings offset.");
            }

            CffIndex charStrings = ReadIndex(cffData, charStringsOffset);
            if (charStrings.Count != expectedGlyphCount) {
                return PdfCffCharStringSubset.Unsupported(
                    cffData,
                    "The CFF CharStrings count does not match the OpenType glyph count.");
            }

            var retainedGlyphIds = new HashSet<int> { 0 };
            for (int index = 0; index < usedGlyphIds.Count; index++) {
                int glyphId = usedGlyphIds[index];
                if (glyphId < 0 || glyphId >= charStrings.Count) {
                    return PdfCffCharStringSubset.Unsupported(cffData, "A used glyph identifier is outside the CFF CharStrings range.");
                }

                retainedGlyphIds.Add(glyphId);
            }

            int prunedGlyphCount = charStrings.Count - retainedGlyphIds.Count;
            if (prunedGlyphCount <= 0) {
                return PdfCffCharStringSubset.Unchanged(cffData, charStrings.Count, charStrings.DataLength);
            }

            byte[][] programs = new byte[charStrings.Count][];
            int subsetProgramBytes = 0;
            for (int glyphId = 0; glyphId < charStrings.Count; glyphId++) {
                byte[] program = retainedGlyphIds.Contains(glyphId)
                    ? CopyObject(cffData, charStrings.GetObject(glyphId))
                    : new[] { Type2EndChar };
                programs[glyphId] = program;
                subsetProgramBytes = checked(subsetProgramBytes + program.Length);
            }

            byte[] subsetIndex = BuildIndex(programs, subsetProgramBytes);
            if (subsetIndex.Length > charStrings.TotalLength) {
                return PdfCffCharStringSubset.Unsupported(cffData, "The rewritten CFF CharStrings INDEX grew beyond its original bounds.");
            }

            byte[] subsetData = cffData.ToArray();
            Array.Clear(subsetData, charStrings.StartOffset, charStrings.TotalLength);
            Buffer.BlockCopy(subsetIndex, 0, subsetData, charStrings.StartOffset, subsetIndex.Length);
            return PdfCffCharStringSubset.Subset(
                subsetData,
                charStrings.Count,
                retainedGlyphIds.Count,
                prunedGlyphCount,
                charStrings.DataLength,
                subsetProgramBytes,
                charStrings.TotalLength,
                subsetIndex.Length);
        } catch (Exception exception) when (IsCffDataException(exception)) {
            return PdfCffCharStringSubset.Unsupported(cffData, exception.Message);
        }
    }

    private static CffIndex ReadIndex(byte[] data, int offset) {
        EnsureRange(data, offset, 2);
        int count = ReadUInt16(data, offset);
        if (count == 0) {
            return new CffIndex(offset, 0, offset + 2, offset + 2, Array.Empty<int>());
        }

        EnsureRange(data, offset + 2, 1);
        int offSize = data[offset + 2];
        if (offSize < 1 || offSize > 4) {
            throw new NotSupportedException("The CFF INDEX offset size is invalid.");
        }

        int offsetsStart = offset + 3;
        int offsetBytes = checked((count + 1) * offSize);
        EnsureRange(data, offsetsStart, offsetBytes);
        var offsets = new int[count + 1];
        for (int index = 0; index <= count; index++) {
            offsets[index] = ReadOffset(data, offsetsStart + index * offSize, offSize);
            if (offsets[index] < 1 || (index > 0 && offsets[index] < offsets[index - 1])) {
                throw new NotSupportedException("The CFF INDEX offsets are invalid.");
            }
        }

        int dataStart = checked(offsetsStart + offsetBytes);
        int dataLength = offsets[count] - 1;
        EnsureRange(data, dataStart, dataLength);
        return new CffIndex(offset, count, dataStart, checked(dataStart + dataLength), offsets);
    }

    private static int FindTopDictionaryInteger(byte[] data, CffObject dictionary, int targetOperator) {
        var operands = new List<double>();
        int offset = dictionary.Offset;
        int end = checked(dictionary.Offset + dictionary.Length);
        while (offset < end) {
            int value = data[offset++];
            if (value >= 32 || value == 28 || value == 29 || value == 30 || value == 255) {
                operands.Add(ReadDictionaryNumber(data, ref offset, end, value));
                continue;
            }

            int dictionaryOperator = value;
            if (value == 12) {
                if (offset >= end) {
                    throw new NotSupportedException("The CFF Top DICT contains a truncated escaped operator.");
                }

                dictionaryOperator = 1200 + data[offset++];
            }

            if (dictionaryOperator == targetOperator) {
                if (operands.Count == 0) {
                    throw new NotSupportedException("The CFF Top DICT CharStrings operator has no operand.");
                }

                double operand = operands[operands.Count - 1];
                if (operand < 0 || operand > int.MaxValue || Math.Abs(operand - Math.Round(operand)) > 0.00001D) {
                    throw new NotSupportedException("The CFF Top DICT CharStrings offset is not an integer.");
                }

                return checked((int)Math.Round(operand));
            }

            operands.Clear();
        }

        return -1;
    }

    private static double ReadDictionaryNumber(byte[] data, ref int offset, int end, int firstByte) {
        if (firstByte >= 32 && firstByte <= 246) {
            return firstByte - 139;
        }

        if (firstByte >= 247 && firstByte <= 250) {
            EnsureDictionaryRange(offset, 1, end);
            return (firstByte - 247) * 256 + data[offset++] + 108;
        }

        if (firstByte >= 251 && firstByte <= 254) {
            EnsureDictionaryRange(offset, 1, end);
            return -(firstByte - 251) * 256 - data[offset++] - 108;
        }

        if (firstByte == 28) {
            EnsureDictionaryRange(offset, 2, end);
            int result = unchecked((short)((data[offset] << 8) | data[offset + 1]));
            offset += 2;
            return result;
        }

        if (firstByte == 29) {
            EnsureDictionaryRange(offset, 4, end);
            int result = unchecked((int)(((uint)data[offset] << 24) |
                ((uint)data[offset + 1] << 16) |
                ((uint)data[offset + 2] << 8) |
                data[offset + 3]));
            offset += 4;
            return result;
        }

        if (firstByte == 255) {
            EnsureDictionaryRange(offset, 4, end);
            int fixedValue = unchecked((int)(((uint)data[offset] << 24) |
                ((uint)data[offset + 1] << 16) |
                ((uint)data[offset + 2] << 8) |
                data[offset + 3]));
            offset += 4;
            return fixedValue / 65536D;
        }

        if (firstByte == 30) {
            var text = new StringBuilder();
            bool complete = false;
            while (offset < end && !complete) {
                int packed = data[offset++];
                AppendRealNibble(text, packed >> 4, ref complete);
                if (!complete) {
                    AppendRealNibble(text, packed & 0x0F, ref complete);
                }
            }

            if (!complete || !double.TryParse(text.ToString(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double real)) {
                throw new NotSupportedException("The CFF Top DICT contains an invalid real operand.");
            }

            return real;
        }

        throw new NotSupportedException("The CFF Top DICT contains an unsupported operand encoding.");
    }

    private static void AppendRealNibble(StringBuilder text, int nibble, ref bool complete) {
        switch (nibble) {
            case 0x0A:
                text.Append('.');
                break;
            case 0x0B:
                text.Append('E');
                break;
            case 0x0C:
                text.Append("E-");
                break;
            case 0x0E:
                text.Append('-');
                break;
            case 0x0F:
                complete = true;
                break;
            case 0x0D:
                throw new NotSupportedException("The CFF Top DICT real operand contains a reserved nibble.");
            default:
                text.Append((char)('0' + nibble));
                break;
        }
    }

    private static byte[] BuildIndex(byte[][] objects, int dataLength) {
        int maximumOffset = checked(dataLength + 1);
        int offSize = maximumOffset <= byte.MaxValue
            ? 1
            : maximumOffset <= ushort.MaxValue
                ? 2
                : maximumOffset <= 0xFFFFFF
                    ? 3
                    : 4;
        int headerLength = checked(3 + (objects.Length + 1) * offSize);
        byte[] result = new byte[checked(headerLength + dataLength)];
        WriteUInt16(result, 0, checked((ushort)objects.Length));
        result[2] = (byte)offSize;

        int relativeOffset = 1;
        int dataOffset = headerLength;
        for (int index = 0; index < objects.Length; index++) {
            WriteOffset(result, 3 + index * offSize, offSize, relativeOffset);
            byte[] value = objects[index];
            Buffer.BlockCopy(value, 0, result, dataOffset, value.Length);
            dataOffset += value.Length;
            relativeOffset = checked(relativeOffset + value.Length);
        }

        WriteOffset(result, 3 + objects.Length * offSize, offSize, relativeOffset);
        return result;
    }

    private static byte[] CopyObject(byte[] data, CffObject value) {
        var result = new byte[value.Length];
        Buffer.BlockCopy(data, value.Offset, result, 0, value.Length);
        return result;
    }

    private static int ReadOffset(byte[] data, int offset, int size) {
        int value = 0;
        for (int index = 0; index < size; index++) {
            value = checked((value << 8) | data[offset + index]);
        }

        return value;
    }

    private static void WriteOffset(byte[] data, int offset, int size, int value) {
        for (int index = size - 1; index >= 0; index--) {
            data[offset + index] = (byte)(value & 0xFF);
            value >>= 8;
        }

        if (value != 0) {
            throw new NotSupportedException("The CFF INDEX offset exceeds its selected offset size.");
        }
    }

    private static int ReadUInt16(byte[] data, int offset) =>
        (data[offset] << 8) | data[offset + 1];

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)(value & 0xFF);
    }

    private static void EnsureRange(byte[] data, int offset, int length) {
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            throw new NotSupportedException("The CFF data is truncated or invalid.");
        }
    }

    private static void EnsureDictionaryRange(int offset, int length, int end) {
        if (offset < 0 || length < 0 || offset > end - length) {
            throw new NotSupportedException("The CFF Top DICT operand is truncated.");
        }
    }

    private static bool IsCffDataException(Exception exception) =>
        exception is NotSupportedException ||
        exception is ArgumentException ||
        exception is ArithmeticException ||
        exception is IndexOutOfRangeException;

    private readonly struct CffIndex {
        private readonly int[] _offsets;

        internal CffIndex(int startOffset, int count, int dataOffset, int endOffset, int[] offsets) {
            StartOffset = startOffset;
            Count = count;
            DataOffset = dataOffset;
            EndOffset = endOffset;
            _offsets = offsets;
        }

        internal int StartOffset { get; }
        internal int Count { get; }
        internal int DataOffset { get; }
        internal int EndOffset { get; }
        internal int DataLength => EndOffset - DataOffset;
        internal int TotalLength => EndOffset - StartOffset;

        internal CffObject GetObject(int index) {
            if (index < 0 || index >= Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            int start = checked(DataOffset + _offsets[index] - 1);
            int end = checked(DataOffset + _offsets[index + 1] - 1);
            return new CffObject(start, end - start);
        }
    }

    private readonly struct CffObject {
        internal CffObject(int offset, int length) {
            Offset = offset;
            Length = length;
        }

        internal int Offset { get; }
        internal int Length { get; }
    }
}

internal sealed class PdfCffCharStringSubset {
    private PdfCffCharStringSubset(
        byte[] data,
        bool isSubset,
        string? unsupportedReason,
        int glyphCount,
        int retainedGlyphCount,
        int prunedGlyphCount,
        int originalProgramBytes,
        int subsetProgramBytes,
        int originalIndexBytes,
        int subsetIndexBytes) {
        Data = data;
        IsSubset = isSubset;
        UnsupportedReason = unsupportedReason;
        GlyphCount = glyphCount;
        RetainedGlyphCount = retainedGlyphCount;
        PrunedGlyphCount = prunedGlyphCount;
        OriginalProgramBytes = originalProgramBytes;
        SubsetProgramBytes = subsetProgramBytes;
        OriginalIndexBytes = originalIndexBytes;
        SubsetIndexBytes = subsetIndexBytes;
    }

    internal byte[] Data { get; }
    internal bool IsSubset { get; }
    internal string? UnsupportedReason { get; }
    internal int GlyphCount { get; }
    internal int RetainedGlyphCount { get; }
    internal int PrunedGlyphCount { get; }
    internal int OriginalProgramBytes { get; }
    internal int SubsetProgramBytes { get; }
    internal int OriginalIndexBytes { get; }
    internal int SubsetIndexBytes { get; }

    internal static PdfCffCharStringSubset Subset(byte[] data, int glyphCount, int retainedGlyphCount, int prunedGlyphCount, int originalProgramBytes, int subsetProgramBytes, int originalIndexBytes, int subsetIndexBytes) =>
        new(data, true, null, glyphCount, retainedGlyphCount, prunedGlyphCount, originalProgramBytes, subsetProgramBytes, originalIndexBytes, subsetIndexBytes);

    internal static PdfCffCharStringSubset Unchanged(byte[] data, int glyphCount, int programBytes) =>
        new(data.ToArray(), false, null, glyphCount, glyphCount, 0, programBytes, programBytes, 0, 0);

    internal static PdfCffCharStringSubset Unsupported(byte[] data, string reason) =>
        new(data.ToArray(), false, reason, 0, 0, 0, 0, 0, 0, 0);
}
