using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal sealed class OfficeImageMetadataSnapshot {
    internal OfficeImageMetadataKinds Kinds { get; set; }
    internal byte[]? Exif { get; set; }
    internal byte[]? Xmp { get; set; }
    internal byte[]? Icc { get; set; }
    internal bool HasExtendedJpegXmp { get; set; }
    internal bool HasDuplicateStandardJpegXmp { get; set; }
    internal bool ExifContainsResolution { get; set; }
    internal bool HasPhysicalResolution { get; set; }
    internal bool HasUnitlessResolution { get; set; }
    internal double? PhysicalDpiX { get; set; }
    internal double? PhysicalDpiY { get; set; }
}

internal static class OfficeImageMetadataInspector {
    private static readonly byte[] ExifPrefix = { (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0 };
    private static readonly byte[] XmpPrefix = System.Text.Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
    private static readonly byte[] ExtendedXmpPrefix = System.Text.Encoding.ASCII.GetBytes("http://ns.adobe.com/xmp/extension/\0");
    private static readonly byte[] IccPrefix = System.Text.Encoding.ASCII.GetBytes("ICC_PROFILE\0");

    internal static OfficeImageMetadataSnapshot Inspect(byte[] data, OfficeImageFormat format) {
        var snapshot = new OfficeImageMetadataSnapshot();
        if (OfficeImageOrientationNormalizer.TryRead(data, out OfficeImageOrientation orientation) &&
            orientation != OfficeImageOrientation.Normal) snapshot.Kinds |= OfficeImageMetadataKinds.Orientation;
        switch (format) {
            case OfficeImageFormat.Jpeg:
                InspectJpeg(data, snapshot);
                break;
            case OfficeImageFormat.Png:
                InspectPng(data, snapshot);
                break;
            case OfficeImageFormat.Webp:
                InspectWebp(data, snapshot);
                break;
            case OfficeImageFormat.Tiff:
                InspectTiff(data, snapshot);
                break;
            case OfficeImageFormat.Bmp:
                InspectBmp(data, snapshot);
                break;
            case OfficeImageFormat.Gif:
                if (HasGifCommentExtension(data)) snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
                break;
        }
        return snapshot;
    }

    private static void InspectJpeg(byte[] data, OfficeImageMetadataSnapshot snapshot) {
        var icc = new SortedDictionary<int, byte[]>();
        int iccTotal = 0;
        int offset = 2;
        bool inScan = false;
        while (offset < data.Length - 1) {
            if (inScan && data[offset] != 0xFF) {
                offset++;
                continue;
            }
            if (data[offset++] != 0xFF) break;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) break;
            int marker = data[offset++];
            if (inScan) {
                if (marker == 0x00 || marker >= 0xD0 && marker <= 0xD7) continue;
                inScan = false;
            }
            if (marker == 0xD9) break;
            if (marker >= 0xD0 && marker <= 0xD7 || marker == 0x01) continue;
            if (offset > data.Length - 2) break;
            int length = data[offset] << 8 | data[offset + 1];
            if (length < 2 || offset > data.Length - length) break;
            int payload = offset + 2;
            int count = length - 2;
            if (marker == 0xE0 && Matches(data, payload, count, "JFIF\0")) {
                bool physical = count >= 12 && data[payload + 7] >= 1 && data[payload + 7] <= 2;
                MarkResolution(snapshot, physical);
                if (physical) {
                    int densityX = data[payload + 8] << 8 | data[payload + 9];
                    int densityY = data[payload + 10] << 8 | data[payload + 11];
                    double scale = data[payload + 7] == 2 ? 2.54D : 1D;
                    SetPhysicalResolution(snapshot, densityX * scale, densityY * scale, overwrite: true);
                }
            }
            if (marker == 0xE1 && StartsWith(data, payload, count, ExifPrefix)) {
                snapshot.Exif = Slice(data, payload, count);
                InspectExifPayload(snapshot.Exif, snapshot);
            } else if (marker == 0xE1 && StartsWith(data, payload, count, XmpPrefix)) {
                if (snapshot.Xmp != null) snapshot.HasDuplicateStandardJpegXmp = true;
                snapshot.Xmp = Slice(data, payload, count);
                snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            } else if (marker == 0xE1 && StartsWith(data, payload, count, ExtendedXmpPrefix)) {
                snapshot.HasExtendedJpegXmp = true;
                snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            } else if (marker == 0xE2 && StartsWith(data, payload, count, IccPrefix) && count >= IccPrefix.Length + 2) {
                int sequence = data[payload + IccPrefix.Length];
                int total = data[payload + IccPrefix.Length + 1];
                if (sequence >= 1 && total >= 1 && sequence <= total) {
                    iccTotal = total;
                    icc[sequence] = Slice(data, payload + IccPrefix.Length + 2, count - IccPrefix.Length - 2);
                }
            } else if (marker == 0xFE) {
                snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            }
            if (marker == 0xDA) inScan = true;
            offset += length;
        }
        if (iccTotal > 0 && icc.Count == iccTotal) {
            int length = 0;
            for (int index = 1; index <= iccTotal; index++) {
                if (!icc.TryGetValue(index, out byte[]? part)) return;
                length = checked(length + part.Length);
            }
            snapshot.Icc = new byte[length];
            int target = 0;
            for (int index = 1; index <= iccTotal; index++) {
                byte[] part = icc[index];
                Buffer.BlockCopy(part, 0, snapshot.Icc, target, part.Length);
                target += part.Length;
            }
            snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
        }
    }

    private static void InspectPng(byte[] data, OfficeImageMetadataSnapshot snapshot) {
        int offset = 8;
        while (offset <= data.Length - 12) {
            int length = ReadBigEndian(data, offset);
            if (length < 0 || offset > data.Length - 12 - length) break;
            string type = ReadAscii(data, offset + 4, 4);
            if (type == "eXIf") {
                snapshot.Exif = Slice(data, offset + 8, length);
                InspectExifPayload(snapshot.Exif, snapshot);
            }
            else if (type == "iCCP") snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
            else if (type == "pHYs") {
                bool physical = length == 9 && data[offset + 16] == 1;
                MarkResolution(snapshot, physical);
                if (physical) {
                    const double pixelsPerMeterPerDpi = 39.37007874015748D;
                    SetPhysicalResolution(snapshot,
                        ReadUInt32Unsigned(data, offset + 8, little: false) / pixelsPerMeterPerDpi,
                        ReadUInt32Unsigned(data, offset + 12, little: false) / pixelsPerMeterPerDpi,
                        overwrite: true);
                }
            }
            else if ((type == "tEXt" || type == "zTXt" || type == "iTXt") &&
                     HasExactPngTextKeyword(
                         data, offset + 8, length, "XML:com.adobe.xmp")) snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (type == "tEXt" || type == "zTXt" || type == "iTXt") snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            offset = checked(offset + 12 + length);
        }
    }

    private static void InspectWebp(byte[] data, OfficeImageMetadataSnapshot snapshot) {
        int offset = 12;
        while (offset <= data.Length - 8) {
            int length = ReadLittleEndian(data, offset + 4);
            if (length < 0 || offset > data.Length - 8 - length) break;
            string type = ReadAscii(data, offset, 4);
            if (type == "EXIF") {
                snapshot.Exif = Slice(data, offset + 8, length);
                InspectExifPayload(snapshot.Exif, snapshot);
            }
            else if (type == "XMP ") snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (type == "ICCP") snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
            offset = checked(offset + 8 + length + (length & 1));
        }
    }

    private static void InspectTiff(byte[] data, OfficeImageMetadataSnapshot snapshot) {
        if (data.Length < 10) return;
        bool little = data[0] == (byte)'I';
        int ifd = ReadUInt32(data, 4, little);
        if (ifd < 0 || ifd > data.Length - 2) return;
        int count = ReadUInt16(data, ifd, little);
        bool hasResolution = false;
        int resolutionUnit = 2;
        double? resolutionX = null;
        double? resolutionY = null;
        for (int index = 0; index < count; index++) {
            int entry = ifd + 2 + index * 12;
            if (entry > data.Length - 12) return;
            int tag = ReadUInt16(data, entry, little);
            if (tag == 34665) snapshot.Kinds |= OfficeImageMetadataKinds.Exif;
            else if (tag == 700) snapshot.Kinds |= OfficeImageMetadataKinds.Xmp;
            else if (tag == 34675) snapshot.Kinds |= OfficeImageMetadataKinds.Icc;
            else if (tag == 270) snapshot.Kinds |= OfficeImageMetadataKinds.Comments;
            else if (tag == 282 || tag == 283) {
                hasResolution = true;
                if (TryReadRational(data, entry, little, tiffBaseOffset: 0, out double resolution)) {
                    if (tag == 282) resolutionX = resolution;
                    else resolutionY = resolution;
                }
            }
            else if (tag == 296) {
                hasResolution = true;
                if (!TryReadInlineShort(data, entry, little, out resolutionUnit)) resolutionUnit = 1;
            }
        }
        if (hasResolution) {
            bool physical = resolutionUnit == 2 || resolutionUnit == 3;
            MarkResolution(snapshot, physical);
            if (physical && resolutionX.HasValue && resolutionY.HasValue) {
                double scale = resolutionUnit == 3 ? 2.54D : 1D;
                SetPhysicalResolution(snapshot, resolutionX.Value * scale, resolutionY.Value * scale, overwrite: true);
            }
        }
    }

    private static bool HasExactPngTextKeyword(
        byte[] data,
        int payloadOffset,
        int payloadLength,
        string keyword) {
        if (payloadLength <= keyword.Length || payloadOffset < 0 ||
            payloadOffset > data.Length - payloadLength) return false;
        for (int index = 0; index < keyword.Length; index++) {
            if (data[payloadOffset + index] != (byte)keyword[index]) return false;
        }
        return data[payloadOffset + keyword.Length] == 0;
    }

    private static void InspectBmp(byte[] data, OfficeImageMetadataSnapshot snapshot) {
        const int dibHeaderOffset = 14;
        const int minimumInfoHeaderSize = 40;
        const int horizontalPixelsPerMeterOffset = 38;
        const int verticalPixelsPerMeterOffset = 42;
        if (data.Length < verticalPixelsPerMeterOffset + 4 ||
            ReadLittleEndian(data, dibHeaderOffset) < minimumInfoHeaderSize) return;

        int horizontalPixelsPerMeter = ReadLittleEndian(data, horizontalPixelsPerMeterOffset);
        int verticalPixelsPerMeter = ReadLittleEndian(data, verticalPixelsPerMeterOffset);
        if (horizontalPixelsPerMeter > 0 || verticalPixelsPerMeter > 0) {
            MarkResolution(snapshot, isPhysical: true);
            if (horizontalPixelsPerMeter > 0 && verticalPixelsPerMeter > 0) {
                const double pixelsPerMeterPerDpi = 39.37007874015748D;
                SetPhysicalResolution(snapshot,
                    horizontalPixelsPerMeter / pixelsPerMeterPerDpi,
                    verticalPixelsPerMeter / pixelsPerMeterPerDpi,
                    overwrite: true);
            }
        }
    }

    private static void InspectExifPayload(byte[] exif, OfficeImageMetadataSnapshot snapshot) {
        snapshot.Kinds |= OfficeImageMetadataKinds.Exif;
        if (OfficeImageOrientationNormalizer.TryReadExifOrientationPayload(exif, out OfficeImageOrientation orientation) &&
            orientation != OfficeImageOrientation.Normal) {
            snapshot.Kinds |= OfficeImageMetadataKinds.Orientation;
        }
        int tiffOffset = exif.Length >= 6 && StartsWith(exif, 0, exif.Length, ExifPrefix) ? 6 : 0;
        if (exif.Length - tiffOffset < 10) return;
        bool little = exif[tiffOffset] == (byte)'I' && exif[tiffOffset + 1] == (byte)'I';
        bool big = exif[tiffOffset] == (byte)'M' && exif[tiffOffset + 1] == (byte)'M';
        if (!little && !big) return;
        int ifd = ReadUInt32(exif, tiffOffset + 4, little);
        if (ifd < 0 || ifd > exif.Length - tiffOffset - 2) return;
        int absoluteIfd = tiffOffset + ifd;
        int count = ReadUInt16(exif, absoluteIfd, little);
        bool hasResolution = false;
        int resolutionUnit = 2;
        double? resolutionX = null;
        double? resolutionY = null;
        for (int index = 0; index < count; index++) {
            int entry = absoluteIfd + 2 + index * 12;
            if (entry > exif.Length - 12) return;
            int tag = ReadUInt16(exif, entry, little);
            if (tag == 282 || tag == 283) {
                hasResolution = true;
                if (TryReadRational(exif, entry, little, tiffOffset, out double resolution)) {
                    if (tag == 282) resolutionX = resolution;
                    else resolutionY = resolution;
                }
            } else if (tag == 296) {
                hasResolution = true;
                if (!TryReadInlineShort(exif, entry, little, out resolutionUnit)) resolutionUnit = 1;
            }
        }
        if (hasResolution) {
            snapshot.ExifContainsResolution = true;
            bool physical = resolutionUnit == 2 || resolutionUnit == 3;
            MarkResolution(snapshot, physical);
            if (physical && resolutionX.HasValue && resolutionY.HasValue) {
                double scale = resolutionUnit == 3 ? 2.54D : 1D;
                SetPhysicalResolution(snapshot, resolutionX.Value * scale, resolutionY.Value * scale, overwrite: false);
            }
        }
    }

    private static void MarkResolution(OfficeImageMetadataSnapshot snapshot, bool isPhysical) {
        snapshot.Kinds |= OfficeImageMetadataKinds.Resolution;
        if (isPhysical) snapshot.HasPhysicalResolution = true;
        else snapshot.HasUnitlessResolution = true;
    }

    private static void SetPhysicalResolution(
        OfficeImageMetadataSnapshot snapshot,
        double dpiX,
        double dpiY,
        bool overwrite) {
        if (double.IsNaN(dpiX) || double.IsInfinity(dpiX) ||
            double.IsNaN(dpiY) || double.IsInfinity(dpiY) || dpiX <= 0D || dpiY <= 0D) return;
        if (overwrite || !snapshot.PhysicalDpiX.HasValue) snapshot.PhysicalDpiX = dpiX;
        if (overwrite || !snapshot.PhysicalDpiY.HasValue) snapshot.PhysicalDpiY = dpiY;
    }

    private static bool TryReadRational(
        byte[] data,
        int entry,
        bool little,
        int tiffBaseOffset,
        out double value) {
        value = 0D;
        if (entry < 0 || entry > data.Length - 12 ||
            ReadUInt16(data, entry + 2, little) != 5 ||
            ReadUInt32Unsigned(data, entry + 4, little) != 1U) return false;
        uint relativeOffset = ReadUInt32Unsigned(data, entry + 8, little);
        long absoluteOffset = (long)tiffBaseOffset + relativeOffset;
        if (absoluteOffset < 0 || absoluteOffset > data.Length - 8) return false;
        uint numerator = ReadUInt32Unsigned(data, (int)absoluteOffset, little);
        uint denominator = ReadUInt32Unsigned(data, (int)absoluteOffset + 4, little);
        if (denominator == 0U) return false;
        value = numerator / (double)denominator;
        return !double.IsNaN(value) && !double.IsInfinity(value) && value > 0D;
    }

    private static bool TryReadInlineShort(byte[] data, int entry, bool little, out int value) {
        value = 0;
        if (entry < 0 || entry > data.Length - 12 ||
            ReadUInt16(data, entry + 2, little) != 3 ||
            ReadUInt32(data, entry + 4, little) != 1) return false;
        value = ReadUInt16(data, entry + 8, little);
        return true;
    }

    private static bool StartsWith(byte[] data, int offset, int count, byte[] prefix) {
        if (count < prefix.Length) return false;
        for (int index = 0; index < prefix.Length; index++) if (data[offset + index] != prefix[index]) return false;
        return true;
    }
    private static byte[] Slice(byte[] data, int offset, int count) {
        var result = new byte[count];
        Buffer.BlockCopy(data, offset, result, 0, count);
        return result;
    }
    private static string ReadAscii(byte[] data, int offset, int count) =>
        System.Text.Encoding.ASCII.GetString(data, offset, count);
    private static bool Matches(byte[] data, int offset, int count, string value) =>
        count >= value.Length && ReadAscii(data, offset, value.Length) == value;
    private static bool HasGifCommentExtension(byte[] data) {
        if (data.Length < 14) return false;
        int offset = 13;
        int packed = data[10];
        if ((packed & 0x80) != 0) offset += 3 << ((packed & 7) + 1);
        while (offset < data.Length) {
            int introducer = data[offset++];
            if (introducer == 0x3B) return false;
            if (introducer == 0x21) {
                if (offset >= data.Length) return false;
                int label = data[offset++];
                if (label == 0xFE) return true;
                if (!SkipGifSubBlocks(data, ref offset)) return false;
                continue;
            }
            if (introducer != 0x2C || offset > data.Length - 9) return false;
            int descriptor = data[offset + 8];
            offset += 9;
            if ((descriptor & 0x80) != 0) offset += 3 << ((descriptor & 7) + 1);
            if (offset >= data.Length) return false;
            offset++;
            if (!SkipGifSubBlocks(data, ref offset)) return false;
        }
        return false;
    }
    private static bool SkipGifSubBlocks(byte[] data, ref int offset) {
        while (offset < data.Length) {
            int length = data[offset++];
            if (length == 0) return true;
            if (offset > data.Length - length) return false;
            offset += length;
        }
        return false;
    }
    private static int ReadBigEndian(byte[] data, int offset) =>
        data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3];
    private static int ReadLittleEndian(byte[] data, int offset) =>
        data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24;
    private static int ReadUInt16(byte[] data, int offset, bool little) => little
        ? data[offset] | data[offset + 1] << 8
        : data[offset] << 8 | data[offset + 1];
    private static int ReadUInt32(byte[] data, int offset, bool little) => little
        ? ReadLittleEndian(data, offset)
        : ReadBigEndian(data, offset);
    private static uint ReadUInt32Unsigned(byte[] data, int offset, bool little) => little
        ? (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24)
        : (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);
}
