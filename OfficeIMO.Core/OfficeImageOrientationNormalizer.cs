using System;

namespace OfficeIMO.Drawing;

/// <summary>EXIF/TIFF image orientation values used by managed image normalization.</summary>
public enum OfficeImageOrientation {
    /// <summary>Rows run top to bottom and columns run left to right.</summary>
    Normal = 1,
    /// <summary>The image is mirrored horizontally.</summary>
    MirrorHorizontal = 2,
    /// <summary>The image is rotated 180 degrees.</summary>
    Rotate180 = 3,
    /// <summary>The image is mirrored vertically.</summary>
    MirrorVertical = 4,
    /// <summary>The image is transposed across its top-left to bottom-right axis.</summary>
    Transpose = 5,
    /// <summary>The image is rotated 90 degrees clockwise.</summary>
    Rotate90Clockwise = 6,
    /// <summary>The image is transposed across its top-right to bottom-left axis.</summary>
    Transverse = 7,
    /// <summary>The image is rotated 90 degrees counter-clockwise.</summary>
    Rotate90CounterClockwise = 8
}

/// <summary>Reads and normalizes embedded JPEG EXIF and classic-TIFF orientation without platform imaging dependencies.</summary>
public static class OfficeImageOrientationNormalizer {
    /// <summary>Attempts to read an explicit non-default or default orientation from JPEG EXIF or classic TIFF bytes.</summary>
    public static bool TryRead(byte[]? imageBytes, out OfficeImageOrientation orientation) {
        orientation = OfficeImageOrientation.Normal;
        if (imageBytes == null || imageBytes.Length < 8) return false;
        if (imageBytes[0] == 0xFF && imageBytes[1] == 0xD8) {
            return TryReadJpeg(imageBytes, out orientation);
        }
        return TryReadTiffOrientation(new OfficeByteView(imageBytes), out orientation);
    }

    /// <summary>
    /// Converts an explicitly rotated or mirrored JPEG/TIFF source to orientation-neutral PNG bytes.
    /// Returns false when orientation is absent/default or the managed decoder cannot normalize the payload.
    /// </summary>
    public static bool TryNormalizeToPng(
        byte[]? imageBytes,
        out byte[] normalizedPng,
        out OfficeImageInfo? normalizedInfo) {
        return TryNormalizeToPng(imageBytes, applyEmbeddedOrientation: true, out normalizedPng, out normalizedInfo);
    }

    /// <summary>
    /// Converts an explicitly oriented JPEG/TIFF source to PNG while either applying or ignoring its metadata.
    /// This supports static renderers that implement CSS <c>image-orientation</c> consistently across outputs.
    /// </summary>
    public static bool TryNormalizeToPng(
        byte[]? imageBytes,
        bool applyEmbeddedOrientation,
        out byte[] normalizedPng,
        out OfficeImageInfo? normalizedInfo) {
        normalizedPng = Array.Empty<byte>();
        normalizedInfo = null;
        if (!TryRead(imageBytes, out OfficeImageOrientation orientation)
            || orientation == OfficeImageOrientation.Normal) {
            return false;
        }
        byte[] source = imageBytes!;
        if (!applyEmbeddedOrientation) {
            source = (byte[])imageBytes!.Clone();
            if (!TryNeutralizeOrientation(source)) return false;
        }
        if (!OfficeImageReader.TryIdentify(source, null, out OfficeImageInfo sourceInfo)
            || !OfficeImagePngConverter.TryConvertToPng(source, out normalizedPng)) {
            normalizedPng = Array.Empty<byte>();
            return false;
        }
        if (applyEmbeddedOrientation && SwapsPhysicalAxes(orientation)) {
            if (!OfficePngReader.TryDecode(normalizedPng, out OfficeRasterImage? normalizedRaster) || normalizedRaster == null) {
                normalizedPng = Array.Empty<byte>();
                return false;
            }
            normalizedPng = OfficePngWriter.Encode(normalizedRaster, new OfficePngEncodeOptions {
                DpiX = sourceInfo.DpiY,
                DpiY = sourceInfo.DpiX
            });
        }
        if (!OfficeImageReader.TryIdentify(normalizedPng, null, out OfficeImageInfo identified)) {
            normalizedPng = Array.Empty<byte>();
            return false;
        }
        normalizedInfo = identified;
        return true;
    }

    private static bool SwapsPhysicalAxes(OfficeImageOrientation orientation) =>
        orientation == OfficeImageOrientation.Transpose
        || orientation == OfficeImageOrientation.Rotate90Clockwise
        || orientation == OfficeImageOrientation.Transverse
        || orientation == OfficeImageOrientation.Rotate90CounterClockwise;

    private static bool TryNeutralizeOrientation(byte[] data) {
        if (data.Length < 8) return false;
        if (data[0] != 0xFF || data[1] != 0xD8) {
            return TryWriteTiffOrientation(data, 0, data.Length);
        }
        int offset = 2;
        while (offset < data.Length) {
            if (data[offset] != 0xFF) return false;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) return false;
            byte marker = data[offset++];
            if (marker is 0xD9 or 0xDA) return false;
            if (marker == 0x01 || marker is >= 0xD0 and <= 0xD7) continue;
            if (marker == 0x00 || offset > data.Length - 2) return false;
            int segmentLength = (data[offset] << 8) | data[offset + 1];
            if (segmentLength < 2 || offset > data.Length - segmentLength) return false;
            if (marker == 0xE1
                && segmentLength >= 8
                && data[offset + 2] == (byte)'E' && data[offset + 3] == (byte)'x'
                && data[offset + 4] == (byte)'i' && data[offset + 5] == (byte)'f'
                && data[offset + 6] == 0 && data[offset + 7] == 0) {
                return TryWriteTiffOrientation(data, offset + 8, segmentLength - 8);
            }
            offset += segmentLength;
        }
        return false;
    }

    private static bool TryWriteTiffOrientation(byte[] data, int tiffOffset, int tiffLength) {
        if (tiffOffset < 0 || tiffLength < 8 || tiffOffset > data.Length - tiffLength) return false;
        bool littleEndian = data[tiffOffset] == (byte)'I' && data[tiffOffset + 1] == (byte)'I';
        bool bigEndian = data[tiffOffset] == (byte)'M' && data[tiffOffset + 1] == (byte)'M';
        if ((!littleEndian && !bigEndian) || ReadUInt16(data, tiffOffset + 2, littleEndian) != 42) return false;
        uint relativeIfd = ReadUInt32(data, tiffOffset + 4, littleEndian);
        if (relativeIfd > int.MaxValue) return false;
        int ifdOffset = checked(tiffOffset + (int)relativeIfd);
        int tiffEnd = tiffOffset + tiffLength;
        if (ifdOffset < tiffOffset + 8 || ifdOffset > tiffEnd - 2) return false;
        int entryCount = ReadUInt16(data, ifdOffset, littleEndian);
        if ((long)ifdOffset + 2L + (entryCount * 12L) > tiffEnd) return false;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + (index * 12);
            if (ReadUInt16(data, entryOffset, littleEndian) != 274) continue;
            if (ReadUInt16(data, entryOffset + 2, littleEndian) != 3
                || ReadUInt32(data, entryOffset + 4, littleEndian) != 1) return false;
            data[entryOffset + 8] = littleEndian ? (byte)1 : (byte)0;
            data[entryOffset + 9] = littleEndian ? (byte)0 : (byte)1;
            return true;
        }
        return false;
    }

    internal static bool TryReadExifOrientation(OfficeByteView app1, out int orientation) {
        orientation = 1;
        if (app1.Length < 14
            || app1[0] != (byte)'E' || app1[1] != (byte)'x'
            || app1[2] != (byte)'i' || app1[3] != (byte)'f'
            || app1[4] != 0 || app1[5] != 0
            || !TryReadTiffOrientation(app1.Slice(6), out OfficeImageOrientation parsed)) {
            return false;
        }
        orientation = (int)parsed;
        return true;
    }

    private static bool TryReadJpeg(byte[] data, out OfficeImageOrientation orientation) {
        orientation = OfficeImageOrientation.Normal;
        int offset = 2;
        while (offset < data.Length) {
            if (data[offset] != 0xFF) return false;
            while (offset < data.Length && data[offset] == 0xFF) offset++;
            if (offset >= data.Length) return false;
            byte marker = data[offset++];
            if (marker is 0xD9 or 0xDA) return false;
            if (marker == 0x01 || marker is >= 0xD0 and <= 0xD7) continue;
            if (marker == 0x00 || offset > data.Length - 2) return false;
            int segmentLength = (data[offset] << 8) | data[offset + 1];
            if (segmentLength < 2 || offset > data.Length - segmentLength) return false;
            if (marker == 0xE1
                && TryReadExifOrientation(new OfficeByteView(data).Slice(offset + 2, segmentLength - 2), out int parsed)) {
                orientation = (OfficeImageOrientation)parsed;
                return true;
            }
            offset += segmentLength;
        }
        return false;
    }

    private static bool TryReadTiffOrientation(OfficeByteView tiff, out OfficeImageOrientation orientation) {
        orientation = OfficeImageOrientation.Normal;
        if (tiff.Length < 8) return false;
        bool littleEndian = tiff[0] == (byte)'I' && tiff[1] == (byte)'I';
        bool bigEndian = tiff[0] == (byte)'M' && tiff[1] == (byte)'M';
        if ((!littleEndian && !bigEndian) || ReadUInt16(tiff, 2, littleEndian) != 42) return false;
        uint ifdOffsetValue = ReadUInt32(tiff, 4, littleEndian);
        if (ifdOffsetValue > int.MaxValue) return false;
        int ifdOffset = (int)ifdOffsetValue;
        if (ifdOffset < 8 || ifdOffset > tiff.Length - 2) return false;
        int entryCount = ReadUInt16(tiff, ifdOffset, littleEndian);
        if ((long)ifdOffset + 2L + (entryCount * 12L) > tiff.Length) return false;
        for (int index = 0; index < entryCount; index++) {
            int entryOffset = ifdOffset + 2 + (index * 12);
            if (ReadUInt16(tiff, entryOffset, littleEndian) != 274) continue;
            if (ReadUInt16(tiff, entryOffset + 2, littleEndian) != 3
                || ReadUInt32(tiff, entryOffset + 4, littleEndian) != 1) return false;
            int value = ReadUInt16(tiff, entryOffset + 8, littleEndian);
            if (value < 1 || value > 8) return false;
            orientation = (OfficeImageOrientation)value;
            return true;
        }
        return false;
    }

    private static ushort ReadUInt16(OfficeByteView data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 2) return 0;
        return littleEndian
            ? (ushort)(data[offset] | (data[offset + 1] << 8))
            : (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static uint ReadUInt32(OfficeByteView data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 4) return 0;
        return littleEndian
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
    }

    private static ushort ReadUInt16(byte[] data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 2) return 0;
        return littleEndian
            ? (ushort)(data[offset] | (data[offset + 1] << 8))
            : (ushort)((data[offset] << 8) | data[offset + 1]);
    }

    private static uint ReadUInt32(byte[] data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 4) return 0;
        return littleEndian
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
    }
}
