using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Applies bounded semantic rewrites to copied Exif TIFF metadata.</summary>
internal static class OfficeExifMetadataEditor {
    internal static bool TryRewritePhysicalResolution(
        byte[] exif,
        double dpiX,
        double dpiY,
        out byte[] rewritten) {
        if (!OfficeImageOrientationNormalizer.TryNeutralizeExifOrientation(exif, out rewritten)) return false;
        int tiffOffset = HasExifPrefix(rewritten) ? 6 : 0;
        int tiffLength = rewritten.Length - tiffOffset;
        if (!OfficeTiffStructureValidator.TryValidateExif(rewritten, tiffOffset, tiffLength)) return false;

        bool littleEndian = rewritten[tiffOffset] == (byte)'I';
        uint relativeIfd = ReadUInt32(rewritten, tiffOffset + 4, littleEndian);
        if (relativeIfd > int.MaxValue) return false;
        int ifdOffset = checked(tiffOffset + (int)relativeIfd);
        int entryCount = ReadUInt16(rewritten, ifdOffset, littleEndian);
        bool sawResolutionField = false;
        int xResolutionOffset = -1;
        int xResolutionEntryOffset = -1;
        int yResolutionOffset = -1;
        int yResolutionEntryOffset = -1;
        int resolutionUnitOffset = -1;
        int resolutionUnitEntryOffset = -1;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = checked(ifdOffset + 2 + index * 12);
            int tag = ReadUInt16(rewritten, entryOffset, littleEndian);
            if (tag is 282 or 283) {
                if (ReadUInt16(rewritten, entryOffset + 2, littleEndian) != 5 ||
                    ReadUInt32(rewritten, entryOffset + 4, littleEndian) != 1U) return false;
                uint relativeValue = ReadUInt32(rewritten, entryOffset + 8, littleEndian);
                if (relativeValue > int.MaxValue) return false;
                int valueOffset = checked(tiffOffset + (int)relativeValue);
                if (tag == 282) {
                    if (xResolutionOffset >= 0) return false;
                    xResolutionOffset = valueOffset;
                    xResolutionEntryOffset = entryOffset;
                } else {
                    if (yResolutionOffset >= 0) return false;
                    yResolutionOffset = valueOffset;
                    yResolutionEntryOffset = entryOffset;
                }
                sawResolutionField = true;
            } else if (tag == 296) {
                if (ReadUInt16(rewritten, entryOffset + 2, littleEndian) != 3 ||
                    ReadUInt32(rewritten, entryOffset + 4, littleEndian) != 1U) return false;
                if (resolutionUnitOffset >= 0) return false;
                resolutionUnitOffset = entryOffset + 8;
                resolutionUnitEntryOffset = entryOffset;
                sawResolutionField = true;
            }
        }
        var writableRanges = new List<int>(9);
        if (xResolutionOffset >= 0) {
            writableRanges.Add(xResolutionOffset);
            writableRanges.Add(8);
            writableRanges.Add(xResolutionEntryOffset);
        }
        if (yResolutionOffset >= 0) {
            writableRanges.Add(yResolutionOffset);
            writableRanges.Add(8);
            writableRanges.Add(yResolutionEntryOffset);
        }
        if (resolutionUnitOffset >= 0) {
            writableRanges.Add(resolutionUnitOffset);
            writableRanges.Add(2);
            writableRanges.Add(resolutionUnitEntryOffset);
        }
        if (!OfficeTiffStructureValidator.TryValidateExclusiveWritableRanges(
                rewritten, tiffOffset, tiffLength, writableRanges.ToArray())) return false;
        if (xResolutionOffset >= 0) WriteRational(rewritten, xResolutionOffset, dpiX, littleEndian);
        if (yResolutionOffset >= 0) WriteRational(rewritten, yResolutionOffset, dpiY, littleEndian);
        if (resolutionUnitOffset >= 0) WriteUInt16(rewritten, resolutionUnitOffset, 2, littleEndian);
        return sawResolutionField;
    }

    private static bool HasExifPrefix(byte[] exif) =>
        exif.Length >= 6 && exif[0] == (byte)'E' && exif[1] == (byte)'x' &&
        exif[2] == (byte)'i' && exif[3] == (byte)'f' && exif[4] == 0 && exif[5] == 0;

    private static int ReadUInt16(byte[] bytes, int offset, bool littleEndian) =>
        littleEndian
            ? bytes[offset] | bytes[offset + 1] << 8
            : bytes[offset] << 8 | bytes[offset + 1];

    private static uint ReadUInt32(byte[] bytes, int offset, bool littleEndian) =>
        littleEndian
            ? (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24)
            : (uint)(bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3]);

    private static void WriteRational(byte[] bytes, int offset, double value, bool littleEndian) {
        const uint denominator = 1000U;
        uint numerator = checked((uint)Math.Round(value * denominator));
        WriteUInt32(bytes, offset, numerator, littleEndian);
        WriteUInt32(bytes, offset + 4, denominator, littleEndian);
    }

    private static void WriteUInt16(byte[] bytes, int offset, int value, bool littleEndian) {
        if (littleEndian) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        } else {
            bytes[offset] = (byte)(value >> 8);
            bytes[offset + 1] = (byte)value;
        }
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value, bool littleEndian) {
        if (littleEndian) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        } else {
            bytes[offset] = (byte)(value >> 24);
            bytes[offset + 1] = (byte)(value >> 16);
            bytes[offset + 2] = (byte)(value >> 8);
            bytes[offset + 3] = (byte)value;
        }
    }
}
