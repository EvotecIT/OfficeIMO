using System;

namespace OfficeIMO.Drawing;

/// <summary>Validates linked and embedded color-profile ranges shared by BMP and ICO DIB payloads.</summary>
internal static class OfficeBitmapV5ProfileValidator {
    private const int BitmapV5HeaderSize = 124;
    private const uint ProfileLinked = 0x4C494E4B;
    private const uint ProfileEmbedded = 0x4D424544;

    internal static bool TryValidate(
        byte[] bytes,
        int dibOffset,
        int dibHeaderSize,
        long pixelOffset,
        long pixelLength,
        int payloadEnd,
        out int profileOffset,
        out int profileSize) {
        profileOffset = 0;
        profileSize = 0;
        if (dibHeaderSize != BitmapV5HeaderSize) return true;
        if (dibOffset < 0 || pixelOffset < 0 || pixelLength < 0 ||
            payloadEnd < 0 || payloadEnd > bytes.Length ||
            dibOffset > payloadEnd - BitmapV5HeaderSize ||
            pixelOffset > payloadEnd || pixelLength > payloadEnd - pixelOffset) return false;

        uint colorSpaceType = ReadUInt32LittleEndian(bytes, dibOffset + 56);
        if (colorSpaceType != ProfileLinked && colorSpaceType != ProfileEmbedded) return true;

        uint relativeProfileOffset = ReadUInt32LittleEndian(bytes, dibOffset + 112);
        uint declaredProfileSize = ReadUInt32LittleEndian(bytes, dibOffset + 116);
        if (relativeProfileOffset < BitmapV5HeaderSize || relativeProfileOffset > int.MaxValue ||
            declaredProfileSize == 0 || declaredProfileSize > int.MaxValue) return false;

        long absoluteProfileOffset = (long)dibOffset + relativeProfileOffset;
        long absoluteProfileEnd = absoluteProfileOffset + declaredProfileSize;
        long pixelEnd = pixelOffset + pixelLength;
        if (absoluteProfileOffset > int.MaxValue || absoluteProfileEnd > payloadEnd ||
            absoluteProfileOffset < pixelEnd && absoluteProfileEnd > pixelOffset) return false;

        profileOffset = (int)absoluteProfileOffset;
        profileSize = (int)declaredProfileSize;
        if (colorSpaceType == ProfileEmbedded) {
            return OfficeIccProfileValidator.TryValidate(bytes, profileOffset, profileSize);
        }

        int profileEnd = profileOffset + profileSize;
        if (bytes[profileEnd - 1] != 0) return false;
        for (int index = profileOffset; index < profileEnd - 1; index++) {
            if (bytes[index] == 0) return false;
        }
        return true;
    }

    private static uint ReadUInt32LittleEndian(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
}
