using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int MaximumWebpExifBytes = 1024 * 1024;

    private static bool TryReadWebp(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 20 ||
            GetAscii(data, 0, 4) != "RIFF" ||
            GetAscii(data, 8, 4) != "WEBP") {
            return false;
        }

        long containerLength = 8L + ReadUInt32LittleEndian(data, 4);
        if (containerLength != data.LongLength) return false;

        int width = 0;
        int height = 0;
        int imageWidth = 0;
        int imageHeight = 0;
        bool extended = false;
        bool hasImage = false;
        byte extendedFlags = 0;
        int exifOffset = 0;
        int exifLength = 0;
        int offset = 12;
        while (offset < data.Length) {
            if (offset > data.Length - 8) return false;
            string chunkType = GetAscii(data, offset, 4);
            uint declaredChunkSize = ReadUInt32LittleEndian(data, offset + 4);
            if (declaredChunkSize > int.MaxValue) return false;

            int chunkSize = (int)declaredChunkSize;
            int chunkDataOffset = checked(offset + 8);
            long chunkDataEnd = (long)chunkDataOffset + chunkSize;
            long paddedChunkEnd = chunkDataEnd + (chunkSize & 1);
            if (chunkDataEnd > containerLength ||
                paddedChunkEnd > containerLength ||
                (chunkSize & 1) != 0 && data[(int)chunkDataEnd] != 0) {
                return false;
            }

            if (chunkType == "VP8X") {
                if (offset != 12 || extended || chunkSize != 10) return false;
                extended = true;
                extendedFlags = data[chunkDataOffset];
                if ((extendedFlags & 0xC1) != 0 ||
                    data[chunkDataOffset + 1] != 0 ||
                    data[chunkDataOffset + 2] != 0 ||
                    data[chunkDataOffset + 3] != 0) {
                    return false;
                }

                width = 1 + ReadUInt24LittleEndian(data, chunkDataOffset + 4);
                height = 1 + ReadUInt24LittleEndian(data, chunkDataOffset + 7);
            } else if (chunkType == "VP8L") {
                if (hasImage ||
                    chunkSize < 5 ||
                    data[chunkDataOffset] != 0x2F) {
                    return false;
                }

                imageWidth = 1 + data[chunkDataOffset + 1] + ((data[chunkDataOffset + 2] & 0x3F) << 8);
                imageHeight = 1 + ((data[chunkDataOffset + 2] & 0xC0) >> 6) +
                              (data[chunkDataOffset + 3] << 2) +
                              ((data[chunkDataOffset + 4] & 0x0F) << 10);
                hasImage = true;
            } else if (chunkType == "VP8 ") {
                if (hasImage ||
                    chunkSize < 10 ||
                    data[chunkDataOffset + 3] != 0x9D ||
                    data[chunkDataOffset + 4] != 0x01 ||
                    data[chunkDataOffset + 5] != 0x2A) {
                    return false;
                }

                imageWidth = ReadUInt16LittleEndian(data, chunkDataOffset + 6) & 0x3FFF;
                imageHeight = ReadUInt16LittleEndian(data, chunkDataOffset + 8) & 0x3FFF;
                hasImage = true;
            } else if (chunkType == "EXIF" && exifOffset == 0) {
                exifOffset = chunkDataOffset;
                exifLength = chunkSize;
            }

            offset = (int)paddedChunkEnd;
        }

        if (offset != data.Length || !hasImage) return false;
        if (extended) {
            if (width != imageWidth ||
                height != imageHeight ||
                ((extendedFlags & 0x08) != 0) != (exifOffset != 0)) {
                return false;
            }
        } else {
            width = imageWidth;
            height = imageHeight;
        }

        double dpiX = 96D;
        double dpiY = 96D;
        if (exifOffset != 0 &&
            TryReadWebpExif(data, exifOffset, exifLength, width, height, out OfficeImageInfo exifInfo)) {
            dpiX = exifInfo.DpiX;
            dpiY = exifInfo.DpiY;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Webp, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool TryReadWebpExif(
        byte[] data,
        int offset,
        int length,
        int expectedWidth,
        int expectedHeight,
        out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (length < 8 || length > MaximumWebpExifBytes) return false;
        if (length >= 14 &&
            GetAscii(data, offset, 4) == "Exif" &&
            data[offset + 4] == 0 &&
            data[offset + 5] == 0) {
            offset += 6;
            length -= 6;
        }

        byte[] tiff = new byte[length];
        Buffer.BlockCopy(data, offset, tiff, 0, length);
        return TryReadTiff(tiff, out info) &&
               info.Width == expectedWidth &&
               info.Height == expectedHeight;
    }

    private static uint ReadUInt32LittleEndian(byte[] data, int offset) =>
        offset + 4 <= data.Length
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : 0U;
}
