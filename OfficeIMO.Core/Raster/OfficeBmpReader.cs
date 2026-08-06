using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free BMP decoder for uncompressed 24-bit and 32-bit Windows bitmap images.
/// </summary>
public static class OfficeBmpReader {
    private const int BitmapFileHeaderSize = 14;
    private const int BitmapInfoHeaderSize = 40;
    private const int BiRgbCompression = 0;

    /// <summary>
    /// Attempts to decode an uncompressed BMP image into an RGBA raster buffer.
    /// </summary>
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) {
        image = null;
        try {
            if (bytes == null || bytes.Length < BitmapFileHeaderSize + BitmapInfoHeaderSize) {
                return false;
            }
            OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "BMP payload exceeds size limits.");

            if (bytes[0] != (byte)'B' || bytes[1] != (byte)'M') {
                return false;
            }

            int pixelOffset = ReadInt32LittleEndian(bytes, 10);
            int dibHeaderSize = ReadInt32LittleEndian(bytes, 14);
            if (dibHeaderSize < BitmapInfoHeaderSize || pixelOffset < BitmapFileHeaderSize + dibHeaderSize || pixelOffset >= bytes.Length) {
                return false;
            }

            int width = ReadInt32LittleEndian(bytes, 18);
            int signedHeight = ReadInt32LittleEndian(bytes, 22);
            int planes = ReadUInt16LittleEndian(bytes, 26);
            int bitsPerPixel = ReadUInt16LittleEndian(bytes, 28);
            int compression = ReadInt32LittleEndian(bytes, 30);
            if (width <= 0 || signedHeight == 0 || planes != 1 || compression != BiRgbCompression ||
                (bitsPerPixel != 24 && bitsPerPixel != 32)) {
                return false;
            }

            int height = Math.Abs(signedHeight);
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) return false;
            bool topDown = signedHeight < 0;
            int rowStride = checked(((width * bitsPerPixel) + 31) / 32 * 4);
            if (pixelOffset + ((long)rowStride * height) > bytes.Length) {
                return false;
            }

            OfficeRasterImage result = new OfficeRasterImage(width, height);
            int bytesPerPixel = bitsPerPixel / 8;
            bool hasAlphaChannel = bitsPerPixel == 32 && HasNonZeroAlpha(bytes, pixelOffset, width, height, rowStride);
            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
                int rowOffset = pixelOffset + (sourceY * rowStride);
                for (int x = 0; x < width; x++) {
                    int pixel = rowOffset + (x * bytesPerPixel);
                    byte blue = bytes[pixel];
                    byte green = bytes[pixel + 1];
                    byte red = bytes[pixel + 2];
                    byte alpha = hasAlphaChannel ? bytes[pixel + 3] : (byte)255;
                    result.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, alpha));
                }
            }

            image = result;
            return true;
        } catch {
            image = null;
            return false;
        }
    }

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8);

    private static bool HasNonZeroAlpha(byte[] bytes, int pixelOffset, int width, int height, int rowStride) {
        for (int y = 0; y < height; y++) {
            int rowOffset = pixelOffset + (y * rowStride);
            for (int x = 0; x < width; x++) {
                if (bytes[rowOffset + (x * 4) + 3] != 0) {
                    return true;
                }
            }
        }

        return false;
    }
}
