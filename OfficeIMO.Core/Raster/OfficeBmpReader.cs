using System;
using System.Threading;

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
    public static bool TryDecode(byte[]? bytes, out OfficeRasterImage? image) =>
        TryDecode(bytes, CancellationToken.None, out image);

    internal static bool TryDecode(byte[]? bytes, CancellationToken cancellationToken, out OfficeRasterImage? image) {
        image = null;
        try {
            if (!TryReadLayout(bytes, cancellationToken, out BmpLayout layout)) return false;
            byte[] source = bytes!;
            if (!IsDecodeWorkingSetWithinLimit(source.LongLength, layout.Width, layout.Height)) return false;

            OfficeRasterImage result = new OfficeRasterImage(layout.Width, layout.Height);
            int bytesPerPixel = layout.BitsPerPixel / 8;
            bool hasAlphaChannel = layout.BitsPerPixel == 32 && HasNonZeroAlpha(
                source, layout.PixelOffset, layout.Width, layout.Height, layout.RowStride, cancellationToken);
            for (int y = 0; y < layout.Height; y++) {
                cancellationToken.ThrowIfCancellationRequested();
                int sourceY = layout.TopDown ? y : layout.Height - 1 - y;
                int rowOffset = layout.PixelOffset + (sourceY * layout.RowStride);
                for (int x = 0; x < layout.Width; x++) {
                    if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    int pixel = rowOffset + (x * bytesPerPixel);
                    byte blue = source[pixel];
                    byte green = source[pixel + 1];
                    byte red = source[pixel + 2];
                    byte alpha = hasAlphaChannel ? source[pixel + 3] : (byte)255;
                    result.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, alpha));
                }
            }

            image = result;
            return true;
        } catch (OperationCanceledException) {
            image = null;
            throw;
        } catch {
            image = null;
            return false;
        }
    }

    internal static bool TryValidatePayload(byte[]? bytes, CancellationToken cancellationToken) {
        try {
            return TryReadLayout(bytes, cancellationToken, out _);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        }
    }

    internal static bool IsDecodeWorkingSetWithinLimit(long encodedBytes, int width, int height) {
        try {
            long rgbaBytes = checked((long)width * height * 4L);
            long peakBytes = checked(encodedBytes + 24L + rgbaBytes + 24L);
            return encodedBytes > 0L && width > 0 && height > 0 &&
                   peakBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool TryReadLayout(
        byte[]? bytes,
        CancellationToken cancellationToken,
        out BmpLayout layout) {
        layout = default;
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes == null || bytes.Length < BitmapFileHeaderSize + BitmapInfoHeaderSize) return false;
        OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "BMP payload exceeds size limits.");
        if (bytes[0] != (byte)'B' || bytes[1] != (byte)'M') return false;

        uint declaredFileSize = ReadUInt32LittleEndian(bytes, 2);
        if (declaredFileSize != bytes.Length ||
            ReadUInt16LittleEndian(bytes, 6) != 0 ||
            ReadUInt16LittleEndian(bytes, 8) != 0) return false;

        int pixelOffset = ReadInt32LittleEndian(bytes, 10);
        int dibHeaderSize = ReadInt32LittleEndian(bytes, 14);
        if (!OfficeDibHeaderLayout.IsSupportedWindowsInfoHeaderSize(dibHeaderSize) ||
            pixelOffset < BitmapFileHeaderSize + dibHeaderSize || pixelOffset >= bytes.Length) return false;

        int width = ReadInt32LittleEndian(bytes, 18);
        int signedHeight = ReadInt32LittleEndian(bytes, 22);
        int planes = ReadUInt16LittleEndian(bytes, 26);
        int bitsPerPixel = ReadUInt16LittleEndian(bytes, 28);
        int compression = ReadInt32LittleEndian(bytes, 30);
        if (width <= 0 || signedHeight == 0 || planes != 1 || compression != BiRgbCompression ||
            (bitsPerPixel != 24 && bitsPerPixel != 32)) return false;

        int height = Math.Abs(signedHeight);
        if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) return false;
        int rowStride = checked(((width * bitsPerPixel) + 31) / 32 * 4);
        long pixelLength = (long)rowStride * height;
        if (pixelOffset + pixelLength > bytes.Length ||
            !OfficeBitmapV5ProfileValidator.TryValidate(
                bytes,
                BitmapFileHeaderSize,
                dibHeaderSize,
                pixelOffset,
                pixelLength,
                bytes.Length,
                out _,
                out _)) return false;

        layout = new BmpLayout(
            pixelOffset, width, height, rowStride, bitsPerPixel, signedHeight < 0);
        return true;
    }

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8);

    private static uint ReadUInt32LittleEndian(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));

    private static bool HasNonZeroAlpha(
        byte[] bytes,
        int pixelOffset,
        int width,
        int height,
        int rowStride,
        CancellationToken cancellationToken) {
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int rowOffset = pixelOffset + (y * rowStride);
            for (int x = 0; x < width; x++) {
                if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (bytes[rowOffset + (x * 4) + 3] != 0) {
                    return true;
                }
            }
        }

        return false;
    }

    private readonly struct BmpLayout {
        internal BmpLayout(
            int pixelOffset,
            int width,
            int height,
            int rowStride,
            int bitsPerPixel,
            bool topDown) {
            PixelOffset = pixelOffset;
            Width = width;
            Height = height;
            RowStride = rowStride;
            BitsPerPixel = bitsPerPixel;
            TopDown = topDown;
        }

        internal int PixelOffset { get; }
        internal int Width { get; }
        internal int Height { get; }
        internal int RowStride { get; }
        internal int BitsPerPixel { get; }
        internal bool TopDown { get; }
    }
}
