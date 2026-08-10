using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Validates bounded classic-TIFF structure used by embedded Exif metadata.</summary>
internal static class OfficeTiffStructureValidator {
    private const int MaximumIfdCount = 1024;
    private const int MaximumEntryCount = 65535;

    /// <summary>Checks byte order, typed value ranges, and all reachable IFD pointer chains.</summary>
    internal static bool TryValidateExif(byte[] bytes, int offset, int count) {
        if (bytes == null || offset < 0 || count < 8 || offset > bytes.Length - count) return false;

        bool littleEndian;
        if (bytes[offset] == (byte)'I' && bytes[offset + 1] == (byte)'I') {
            littleEndian = true;
        } else if (bytes[offset] == (byte)'M' && bytes[offset + 1] == (byte)'M') {
            littleEndian = false;
        } else {
            return false;
        }
        if (ReadUInt16(bytes, offset + 2, littleEndian) != 42) return false;

        uint firstIfd = ReadUInt32(bytes, offset + 4, littleEndian);
        if (firstIfd < 8 || firstIfd > int.MaxValue) return false;
        var pending = new Stack<int>();
        pending.Push((int)firstIfd);
        var scheduled = new HashSet<int> { (int)firstIfd };
        var visited = new HashSet<int>();
        int totalEntries = 0;

        while (pending.Count > 0) {
            int ifdOffset = pending.Pop();
            if (!visited.Add(ifdOffset)) return false;
            if (visited.Count > MaximumIfdCount || ifdOffset < 8 || ifdOffset > count - 6) return false;

            int absoluteIfd = offset + ifdOffset;
            int entryCount = ReadUInt16(bytes, absoluteIfd, littleEndian);
            if (entryCount > MaximumEntryCount - totalEntries) return false;
            totalEntries += entryCount;
            long tableLength = 2L + entryCount * 12L + 4L;
            if (tableLength > count - ifdOffset) return false;

            int entryOffset = absoluteIfd + 2;
            for (int index = 0; index < entryCount; index++, entryOffset += 12) {
                ushort tag = ReadUInt16(bytes, entryOffset, littleEndian);
                ushort type = ReadUInt16(bytes, entryOffset + 2, littleEndian);
                uint valueCount = ReadUInt32(bytes, entryOffset + 4, littleEndian);
                int typeSize = GetTypeSize(type);
                if (typeSize == 0) return false;
                ulong valueSize = (ulong)valueCount * (uint)typeSize;
                if (valueSize > int.MaxValue) return false;

                if (valueSize > 4) {
                    uint valueOffset = ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    if (valueOffset > int.MaxValue || valueOffset > (uint)count ||
                        valueSize > (ulong)(count - (int)valueOffset)) return false;
                }

                if ((tag == 34665 || tag == 34853 || tag == 40965) && type == 4 && valueCount == 1) {
                    uint nestedIfd = ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    if (nestedIfd != 0) {
                        if (nestedIfd > int.MaxValue) return false;
                        if (!TryScheduleIfd((int)nestedIfd, pending, scheduled)) return false;
                    }
                } else if (type == 13 && valueCount > 0) {
                    if (valueCount > MaximumIfdCount) return false;
                    int valuesOffset;
                    if (valueSize <= 4) {
                        valuesOffset = entryOffset + 8;
                    } else {
                        valuesOffset = offset + (int)ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    }
                    for (uint valueIndex = 0; valueIndex < valueCount; valueIndex++) {
                        uint nestedIfd = ReadUInt32(bytes, valuesOffset + (int)valueIndex * 4, littleEndian);
                        if (nestedIfd != 0) {
                            if (nestedIfd > int.MaxValue) return false;
                            if (!TryScheduleIfd((int)nestedIfd, pending, scheduled)) return false;
                        }
                    }
                }
            }

            uint nextIfd = ReadUInt32(bytes, entryOffset, littleEndian);
            if (nextIfd != 0) {
                if (nextIfd > int.MaxValue) return false;
                if (!TryScheduleIfd((int)nextIfd, pending, scheduled)) return false;
            }
        }

        return true;
    }

    private static bool TryScheduleIfd(int offset, Stack<int> pending, HashSet<int> scheduled) {
        if (!scheduled.Add(offset) || scheduled.Count > MaximumIfdCount) return false;
        pending.Push(offset);
        return true;
    }

    private static int GetTypeSize(ushort type) {
        switch (type) {
            case 1:
            case 2:
            case 6:
            case 7:
                return 1;
            case 3:
            case 8:
                return 2;
            case 4:
            case 9:
            case 11:
            case 13:
                return 4;
            case 5:
            case 10:
            case 12:
                return 8;
            default:
                return 0;
        }
    }

    private static ushort ReadUInt16(byte[] bytes, int offset, bool littleEndian) {
        return littleEndian
            ? (ushort)(bytes[offset] | bytes[offset + 1] << 8)
            : (ushort)(bytes[offset] << 8 | bytes[offset + 1]);
    }

    private static uint ReadUInt32(byte[] bytes, int offset, bool littleEndian) {
        if (littleEndian) {
            return (uint)(bytes[offset] |
                          bytes[offset + 1] << 8 |
                          bytes[offset + 2] << 16 |
                          bytes[offset + 3] << 24);
        }
        return (uint)(bytes[offset] << 24 |
                      bytes[offset + 1] << 16 |
                      bytes[offset + 2] << 8 |
                      bytes[offset + 3]);
    }
}
