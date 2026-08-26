using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared cmap platform and encoding classification.</summary>
internal static class OfficeOpenTypeCmap {
    internal static bool IsUnicodeEncoding(int platform, int encoding) =>
        platform == 0 ||
        platform == 3 && (encoding == 1 || encoding == 10);

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

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3];
}
