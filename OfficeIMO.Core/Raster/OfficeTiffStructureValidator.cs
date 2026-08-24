using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Validates bounded classic-TIFF structure used by images and embedded Exif metadata.</summary>
internal static class OfficeTiffStructureValidator {
    private const int MaximumIfdCount = 1024;
    private const int MaximumEntryCount = 65535;

    /// <summary>Checks byte order, typed value ranges, and all reachable IFD pointer chains.</summary>
    internal static bool TryValidate(byte[] bytes, int offset, int count) =>
        TryValidate(bytes, offset, count, CancellationToken.None);

    internal static bool TryValidate(
        byte[] bytes,
        int offset,
        int count,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
                if ((index & 0xFF) == 0) cancellationToken.ThrowIfCancellationRequested();
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
                } else if ((type == 13 || tag == 330 && type == 4) && valueCount > 0) {
                    if (valueCount > MaximumIfdCount) return false;
                    int valuesOffset;
                    if (valueSize <= 4) {
                        valuesOffset = entryOffset + 8;
                    } else {
                        valuesOffset = offset + (int)ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    }
                    for (uint valueIndex = 0; valueIndex < valueCount; valueIndex++) {
                        if ((valueIndex & 0xFFU) == 0U) cancellationToken.ThrowIfCancellationRequested();
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

    /// <summary>
    /// Verifies that each requested writable range belongs to exactly one reachable TIFF entry
    /// and does not overlap any other TIFF-owned header, IFD, or value range.
    /// </summary>
    internal static bool TryValidateExclusiveWritableRanges(
        byte[] bytes,
        int offset,
        int count,
        params int[] offsetLengthOwnerTriples) {
        if (offsetLengthOwnerTriples == null || offsetLengthOwnerTriples.Length % 3 != 0 ||
            !TryValidate(bytes, offset, count)) return false;

        bool littleEndian = bytes[offset] == (byte)'I';
        int rangeCount = offsetLengthOwnerTriples.Length / 3;
        var ownerReferences = new int[rangeCount];
        var inlineOwners = new bool[rangeCount];
        for (int range = 0; range < rangeCount; range++) {
            int rangeOffset = offsetLengthOwnerTriples[range * 3];
            int rangeLength = offsetLengthOwnerTriples[range * 3 + 1];
            int ownerEntryOffset = offsetLengthOwnerTriples[range * 3 + 2];
            if (rangeOffset < offset || rangeLength <= 0 || rangeOffset > offset + count - rangeLength ||
                ownerEntryOffset < offset + 8 || ownerEntryOffset > offset + count - 12 ||
                RangesOverlap(rangeOffset, rangeLength, offset, 8)) return false;

            ushort type = ReadUInt16(bytes, ownerEntryOffset + 2, littleEndian);
            uint valueCount = ReadUInt32(bytes, ownerEntryOffset + 4, littleEndian);
            int typeSize = GetTypeSize(type);
            if (typeSize == 0) return false;
            ulong valueSize = (ulong)valueCount * (uint)typeSize;
            if (valueSize == 0 || valueSize > int.MaxValue || rangeLength != (int)valueSize) return false;
            int storageOffset = valueSize <= 4
                ? ownerEntryOffset + 8
                : checked(offset + (int)ReadUInt32(bytes, ownerEntryOffset + 8, littleEndian));
            if (rangeOffset != storageOffset) return false;
            inlineOwners[range] = valueSize <= 4;
        }

        var pending = new Stack<int>();
        var visited = new HashSet<int>();
        pending.Push((int)ReadUInt32(bytes, offset + 4, littleEndian));
        while (pending.Count > 0) {
            int relativeIfd = pending.Pop();
            if (!visited.Add(relativeIfd)) return false;
            int absoluteIfd = offset + relativeIfd;
            int entryCount = ReadUInt16(bytes, absoluteIfd, littleEndian);
            int tableLength = checked(2 + entryCount * 12 + 4);
            for (int range = 0; range < rangeCount; range++) {
                int ownerEntryOffset = offsetLengthOwnerTriples[range * 3 + 2];
                bool ownsEntry = ownerEntryOffset >= absoluteIfd + 2 &&
                    ownerEntryOffset < absoluteIfd + 2 + entryCount * 12 &&
                    (ownerEntryOffset - absoluteIfd - 2) % 12 == 0;
                if (ownsEntry) ownerReferences[range]++;
                if (RangesOverlap(offsetLengthOwnerTriples[range * 3], offsetLengthOwnerTriples[range * 3 + 1],
                        absoluteIfd, tableLength) && !(inlineOwners[range] && ownsEntry)) return false;
            }

            int entryOffset = absoluteIfd + 2;
            for (int index = 0; index < entryCount; index++, entryOffset += 12) {
                ushort tag = ReadUInt16(bytes, entryOffset, littleEndian);
                ushort type = ReadUInt16(bytes, entryOffset + 2, littleEndian);
                uint valueCount = ReadUInt32(bytes, entryOffset + 4, littleEndian);
                int valueLength = checked((int)(valueCount * (uint)GetTypeSize(type)));
                if (valueLength > 4) {
                    int valueOffset = checked(offset + (int)ReadUInt32(bytes, entryOffset + 8, littleEndian));
                    for (int range = 0; range < rangeCount; range++) {
                        int writableOffset = offsetLengthOwnerTriples[range * 3];
                        int writableLength = offsetLengthOwnerTriples[range * 3 + 1];
                        if (!RangesOverlap(writableOffset, writableLength, valueOffset, valueLength)) continue;
                        if (entryOffset != offsetLengthOwnerTriples[range * 3 + 2] ||
                            writableOffset != valueOffset || writableLength != valueLength) return false;
                    }
                }

                if ((tag == 34665 || tag == 34853 || tag == 40965) && type == 4 && valueCount == 1) {
                    uint nestedIfd = ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    if (nestedIfd != 0) pending.Push((int)nestedIfd);
                } else if ((type == 13 || tag == 330 && type == 4) && valueCount > 0) {
                    int valuesOffset = valueLength <= 4
                        ? entryOffset + 8
                        : offset + (int)ReadUInt32(bytes, entryOffset + 8, littleEndian);
                    for (uint valueIndex = 0; valueIndex < valueCount; valueIndex++) {
                        uint nestedIfd = ReadUInt32(bytes, valuesOffset + (int)valueIndex * 4, littleEndian);
                        if (nestedIfd != 0) pending.Push((int)nestedIfd);
                    }
                }
            }

            uint nextIfd = ReadUInt32(bytes, entryOffset, littleEndian);
            if (nextIfd != 0) pending.Push((int)nextIfd);
        }

        for (int range = 0; range < rangeCount; range++) {
            if (ownerReferences[range] != 1) return false;
        }
        return true;
    }

    /// <summary>Checks the classic-TIFF structure carried by an Exif metadata payload.</summary>
    internal static bool TryValidateExif(byte[] bytes, int offset, int count) =>
        TryValidate(bytes, offset, count);

    internal static bool TryValidateExif(
        byte[] bytes,
        int offset,
        int count,
        CancellationToken cancellationToken) =>
        TryValidate(bytes, offset, count, cancellationToken);

    private static bool TryScheduleIfd(int offset, Stack<int> pending, HashSet<int> scheduled) {
        if (!scheduled.Add(offset) || scheduled.Count > MaximumIfdCount) return false;
        pending.Push(offset);
        return true;
    }

    private static bool RangesOverlap(int firstOffset, int firstLength, int secondOffset, int secondLength) =>
        firstOffset < secondOffset + secondLength && secondOffset < firstOffset + firstLength;

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
