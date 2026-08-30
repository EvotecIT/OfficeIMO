using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared cmap platform and encoding classification.</summary>
internal static class OfficeOpenTypeCmap {
    internal const int MaximumSubtables = 64;
    internal const uint MaximumFormat12Groups = 4096;
    private const uint MaximumVariationSelectorRecords = 256;
    private const uint MaximumVariationMappings = 4096;

    internal static bool IsUnicodeEncoding(int platform, int encoding) =>
        platform == 0 ||
        platform == 3 && (encoding == 1 || encoding == 10);

    internal static int ScoreSubtable(int format, int platform, int encoding, bool preferFormat12) {
        int score = format == 12 ? 100 : 50;
        if (preferFormat12 && format == 12) score += 100;
        if (platform == 3 && encoding == 10) score += 20;
        else if (platform == 0) score += 15;
        else if (platform == 3 && encoding == 1) score += 10;
        return score;
    }

    internal static bool HasGlyphs(
        string text,
        Func<int, int> mapGlyph,
        Func<int, int, int> mapVariationSequence) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        for (int index = 0; index < text.Length;) {
            int glyph = ReadMappedGlyph(text, ref index, mapGlyph, mapVariationSequence, out int scalar);
            if (glyph < 0) continue;
            if (glyph == 0 && !(scalar <= char.MaxValue && char.IsWhiteSpace((char)scalar))) return false;
        }
        return true;
    }

    internal static int ReadMappedGlyph(
        string text,
        ref int index,
        Func<int, int> mapGlyph,
        Func<int, int, int> mapVariationSequence,
        out int scalar) {
        scalar = ReadScalar(text, ref index);
        if (IsVariationSelector(scalar)) return 0;

        int followingIndex = index;
        int followingScalar = followingIndex < text.Length
            ? ReadScalar(text, ref followingIndex)
            : -1;
        if (IsVariationSelector(followingScalar)) {
            index = followingIndex;
            return mapVariationSequence(scalar, followingScalar);
        }

        return OfficeTextElements.IsIgnorableFontCoverageScalar(scalar) ? -1 : mapGlyph(scalar);
    }

    internal static int MapVariationSequence(
        byte[] data,
        int cmapOffset,
        int cmapLength,
        int glyphCount,
        int scalar,
        int variationSelector,
        Func<int, int> mapGlyph) {
        if (data == null || cmapOffset < 0 || cmapLength < 4 || cmapOffset > data.Length - cmapLength ||
            glyphCount <= 0 || scalar < 0 || scalar > 0x10FFFF || !IsVariationSelector(variationSelector)) return 0;
        int cmapEnd = cmapOffset + cmapLength;
        int count = ReadUInt16(data, cmapOffset + 2);
        if (count <= 0 || count > MaximumSubtables || cmapLength < 4 + count * 8) return 0;
        for (int index = 0; index < count; index++) {
            int record = cmapOffset + 4 + index * 8;
            int platform = ReadUInt16(data, record);
            int encoding = ReadUInt16(data, record + 2);
            if (platform != 0 || encoding != 5) continue;
            uint relativeValue = ReadUInt32(data, record + 4);
            if (relativeValue > (uint)(cmapLength - 10)) continue;
            int table = cmapOffset + (int)relativeValue;
            if (table < cmapOffset || table > cmapEnd - 10 || ReadUInt16(data, table) != 14) continue;
            int glyph = MapVariationSequenceSubtable(data, table, cmapEnd, glyphCount, scalar, variationSelector, mapGlyph);
            if (glyph != 0) return glyph;
        }
        return 0;
    }

    private static int MapVariationSequenceSubtable(
        byte[] data,
        int table,
        int cmapEnd,
        int glyphCount,
        int scalar,
        int variationSelector,
        Func<int, int> mapGlyph) {
        uint lengthValue = ReadUInt32(data, table + 2);
        uint recordCountValue = ReadUInt32(data, table + 6);
        if (lengthValue > int.MaxValue || recordCountValue > MaximumVariationSelectorRecords) return 0;
        int length = (int)lengthValue;
        int recordCount = (int)recordCountValue;
        if (length < 10 || table > cmapEnd - length || 10L + recordCount * 11L > length) return 0;
        int tableEnd = table + length;

        int matchingRecord = -1;
        int previousSelector = -1;
        for (int index = 0; index < recordCount; index++) {
            int record = table + 10 + index * 11;
            int selector = ReadUInt24(data, record);
            if (!IsVariationSelector(selector) || selector <= previousSelector) return 0;
            if (selector == variationSelector) matchingRecord = record;
            previousSelector = selector;
        }
        if (matchingRecord < 0) return 0;

        uint defaultOffset = ReadUInt32(data, matchingRecord + 3);
        uint nonDefaultOffset = ReadUInt32(data, matchingRecord + 7);
        if (nonDefaultOffset != 0) {
            if (!TryResolveNonDefaultVariation(
                    data,
                    table,
                    tableEnd,
                    nonDefaultOffset,
                    glyphCount,
                    scalar,
                    out int glyph)) return 0;
            if (glyph != 0) return glyph;
        }
        if (defaultOffset == 0 ||
            !TryResolveDefaultVariation(data, table, tableEnd, defaultOffset, scalar, out bool isDefault)) return 0;
        return isDefault ? mapGlyph(scalar) : 0;
    }

    private static bool TryResolveNonDefaultVariation(
        byte[] data,
        int table,
        int tableEnd,
        uint relativeOffset,
        int glyphCount,
        int scalar,
        out int glyph) {
        glyph = 0;
        if (relativeOffset > int.MaxValue) return false;
        int offset = table + (int)relativeOffset;
        if (offset < table || offset > tableEnd - 4) return false;
        uint countValue = ReadUInt32(data, offset);
        if (countValue > MaximumVariationMappings || countValue > (uint)((tableEnd - offset - 4) / 5)) return false;
        int previousScalar = -1;
        for (int index = 0; index < (int)countValue; index++) {
            int mapping = offset + 4 + index * 5;
            int unicodeValue = ReadUInt24(data, mapping);
            int mappedGlyph = ReadUInt16(data, mapping + 3);
            if (unicodeValue > 0x10FFFF || unicodeValue <= previousScalar ||
                mappedGlyph <= 0 || mappedGlyph >= glyphCount) return false;
            if (unicodeValue == scalar) glyph = mappedGlyph;
            previousScalar = unicodeValue;
        }
        return true;
    }

    private static bool TryResolveDefaultVariation(
        byte[] data,
        int table,
        int tableEnd,
        uint relativeOffset,
        int scalar,
        out bool isDefault) {
        isDefault = false;
        if (relativeOffset > int.MaxValue) return false;
        int offset = table + (int)relativeOffset;
        if (offset < table || offset > tableEnd - 4) return false;
        uint countValue = ReadUInt32(data, offset);
        if (countValue > MaximumVariationMappings || countValue > (uint)((tableEnd - offset - 4) / 4)) return false;
        int previousEnd = -1;
        for (int index = 0; index < (int)countValue; index++) {
            int range = offset + 4 + index * 4;
            int start = ReadUInt24(data, range);
            int end = start + data[range + 3];
            if (start <= previousEnd || end > 0x10FFFF) return false;
            if (scalar >= start && scalar <= end) isDefault = true;
            previousEnd = end;
        }
        return true;
    }

    private static bool IsVariationSelector(int scalar) =>
        scalar >= 0xFE00 && scalar <= 0xFE0F || scalar >= 0xE0100 && scalar <= 0xE01EF;

    private static int ReadScalar(string value, ref int index) {
        char first = value[index++];
        if (char.IsHighSurrogate(first) && index < value.Length && char.IsLowSurrogate(value[index])) {
            return char.ConvertToUtf32(first, value[index++]);
        }
        return first;
    }

    internal static HashSet<int> CollectValidFormat12Subtables(
        byte[] data,
        int cmapOffset,
        int cmapLength,
        int maximumSubtables,
        uint maximumGroups) {
        var valid = new HashSet<int>();
        if (data == null || cmapOffset < 0 || cmapLength < 4 || cmapOffset > data.Length - cmapLength) return valid;
        int cmapEnd = cmapOffset + cmapLength;
        int count = ReadUInt16(data, cmapOffset + 2);
        if (count <= 0 || count > maximumSubtables || cmapLength < 4 + count * 8) return valid;
        for (int index = 0; index < count; index++) {
            int record = cmapOffset + 4 + index * 8;
            uint relativeValue = ReadUInt32(data, record + 4);
            if (relativeValue > (uint)(cmapLength - 2)) continue;
            int table = cmapOffset + (int)relativeValue;
            if (table < cmapOffset || table > cmapEnd - 2 || ReadUInt16(data, table) != 12) continue;
            if (IsValidFormat12Subtable(data, table, cmapOffset, cmapEnd, maximumGroups)) valid.Add(table);
        }
        return valid;
    }

    internal static HashSet<int> CollectValidFormat4Subtables(
        byte[] data,
        int cmapOffset,
        int cmapLength,
        int maximumSubtables) {
        var valid = new HashSet<int>();
        if (data == null || cmapOffset < 0 || cmapLength < 4 || cmapOffset > data.Length - cmapLength) return valid;
        int cmapEnd = cmapOffset + cmapLength;
        int count = ReadUInt16(data, cmapOffset + 2);
        if (count <= 0 || count > maximumSubtables || cmapLength < 4 + count * 8) return valid;
        for (int index = 0; index < count; index++) {
            int record = cmapOffset + 4 + index * 8;
            uint relativeValue = ReadUInt32(data, record + 4);
            if (relativeValue > (uint)(cmapLength - 2)) continue;
            int table = cmapOffset + (int)relativeValue;
            if (table < cmapOffset || table > cmapEnd - 2 || ReadUInt16(data, table) != 4) continue;
            if (IsValidFormat4Subtable(data, table, cmapOffset, cmapEnd, maximumSubtables)) valid.Add(table);
        }
        return valid;
    }

    private static bool IsValidFormat4Subtable(
        byte[] data,
        int table,
        int cmapOffset,
        int cmapEnd,
        int maximumSubtables) {
        if (table < cmapOffset || table > cmapEnd - 16) return false;
        int length = ReadUInt16(data, table + 2);
        int segmentCountX2 = ReadUInt16(data, table + 6);
        if (length < 16 || (segmentCountX2 & 1) != 0 || segmentCountX2 == 0 || table > cmapEnd - length) return false;
        int segmentCount = segmentCountX2 / 2;
        if (segmentCount > maximumSubtables * 16) return false;
        int endCodes = table + 14;
        int startCodes = endCodes + segmentCount * 2 + 2;
        int deltas = startCodes + segmentCount * 2;
        int rangeOffsets = deltas + segmentCount * 2;
        int tableEnd = table + length;
        if (rangeOffsets < table || rangeOffsets > tableEnd - segmentCount * 2) return false;

        int previousEnd = -1;
        for (int index = 0; index < segmentCount; index++) {
            int end = ReadUInt16(data, endCodes + index * 2);
            int start = ReadUInt16(data, startCodes + index * 2);
            if (start > end || start <= previousEnd) return false;
            int rangeOffset = ReadUInt16(data, rangeOffsets + index * 2);
            if (rangeOffset != 0) {
                long firstGlyph = (long)rangeOffsets + index * 2 + rangeOffset;
                long lastGlyph = firstGlyph + (long)(end - start) * 2;
                if (firstGlyph < table || lastGlyph > tableEnd - 2L) return false;
            }
            previousEnd = end;
        }
        return ReadUInt16(data, endCodes + (segmentCount - 1) * 2) == 0xFFFF &&
            ReadUInt16(data, startCodes + (segmentCount - 1) * 2) == 0xFFFF;
    }

    private static bool IsValidFormat12Subtable(
        byte[] data,
        int table,
        int cmapOffset,
        int cmapEnd,
        uint maximumGroups) {
        if (table < cmapOffset || table > cmapEnd - 16) return false;
        uint lengthValue = ReadUInt32(data, table + 4);
        uint groupCount = ReadUInt32(data, table + 12);
        if (lengthValue > int.MaxValue || groupCount > maximumGroups) return false;
        int length = (int)lengthValue;
        if (length < 16 || table > cmapEnd - length || 16L + groupCount * 12L > length) return false;

        uint previousEnd = 0;
        for (uint index = 0; index < groupCount; index++) {
            int group = checked(table + 16 + (int)index * 12);
            uint start = ReadUInt32(data, group);
            uint end = ReadUInt32(data, group + 4);
            uint startGlyph = ReadUInt32(data, group + 8);
            if (start > end || end > 0x10FFFFU || index > 0 && start <= previousEnd) return false;
            if ((ulong)startGlyph + end - start > uint.MaxValue) return false;
            previousEnd = end;
        }
        return true;
    }

    private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static int ReadUInt24(byte[] data, int offset) =>
        (data[offset] << 16) | (data[offset + 1] << 8) | data[offset + 2];

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3];
}
