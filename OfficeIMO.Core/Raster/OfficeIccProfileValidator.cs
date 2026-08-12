using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Validates the bounded structural envelope of an embedded ICC profile.</summary>
internal static class OfficeIccProfileValidator {
    private const int HeaderLength = 128;
    private const int TagTableHeaderLength = 4;
    private const int TagEntryLength = 12;
    private const int MaximumTagCount = 65535;

    /// <summary>Checks the declared profile size, signature, reserved header, and every tag range.</summary>
    internal static bool TryValidate(byte[] bytes, int offset, int count) {
        if (bytes == null || offset < 0 || count < HeaderLength + TagTableHeaderLength ||
            (count & 3) != 0 || offset > bytes.Length - count) {
            return false;
        }

        if (ReadUInt32(bytes, offset) != (uint)count ||
            bytes[offset + 36] != (byte)'a' ||
            bytes[offset + 37] != (byte)'c' ||
            bytes[offset + 38] != (byte)'s' ||
            bytes[offset + 39] != (byte)'p') {
            return false;
        }
        for (int index = 100; index < HeaderLength; index++) {
            if (bytes[offset + index] != 0) return false;
        }

        uint declaredTagCount = ReadUInt32(bytes, offset + HeaderLength);
        if (declaredTagCount > MaximumTagCount) return false;
        int tagCount = (int)declaredTagCount;
        long tableEndLong = HeaderLength + TagTableHeaderLength + (long)tagCount * TagEntryLength;
        if (tableEndLong > count) return false;
        int tableEnd = (int)tableEndLong;
        var signatures = new HashSet<uint>();
        var ranges = new List<TagRange>(tagCount);

        for (int index = 0; index < tagCount; index++) {
            int entry = offset + HeaderLength + TagTableHeaderLength + index * TagEntryLength;
            uint signature = ReadUInt32(bytes, entry);
            uint declaredTagOffset = ReadUInt32(bytes, entry + 4);
            uint declaredTagLength = ReadUInt32(bytes, entry + 8);
            if (signature == 0 || !signatures.Add(signature) ||
                declaredTagOffset > int.MaxValue || declaredTagLength < 8 || declaredTagLength > int.MaxValue) {
                return false;
            }

            int tagOffset = (int)declaredTagOffset;
            int tagLength = (int)declaredTagLength;
            long tagEndLong = (long)tagOffset + tagLength;
            long paddedEndLong = (tagEndLong + 3L) & ~3L;
            if ((tagOffset & 3) != 0 || tagOffset < tableEnd || tagEndLong > count || paddedEndLong > count) {
                return false;
            }

            int absoluteTag = offset + tagOffset;
            if (ReadUInt32(bytes, absoluteTag) == 0 ||
                bytes[absoluteTag + 4] != 0 || bytes[absoluteTag + 5] != 0 ||
                bytes[absoluteTag + 6] != 0 || bytes[absoluteTag + 7] != 0) {
                return false;
            }
            for (int pad = (int)tagEndLong; pad < paddedEndLong; pad++) {
                if (bytes[offset + pad] != 0) return false;
            }
            ranges.Add(new TagRange(tagOffset, (int)tagEndLong));
        }

        ranges.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (int index = 1; index < ranges.Count; index++) {
            TagRange previous = ranges[index - 1];
            TagRange current = ranges[index];
            if (current.Start < previous.End &&
                (current.Start != previous.Start || current.End != previous.End)) {
                return false;
            }
        }
        return true;
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        unchecked(((uint)bytes[offset] << 24) |
                  ((uint)bytes[offset + 1] << 16) |
                  ((uint)bytes[offset + 2] << 8) |
                  bytes[offset + 3]);

    private readonly struct TagRange {
        internal TagRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }
}
