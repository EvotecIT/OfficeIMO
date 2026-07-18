using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Defines the zlib compression strategy used when encoding already-filtered PNG scanlines.
/// </summary>
public enum OfficePngCompression {
    /// <summary>
    /// Compress scanlines with the platform deflate implementation.
    /// </summary>
    Optimal,

    /// <summary>
    /// Store scanlines in zlib blocks without deflate compression.
    /// </summary>
    Stored
}

/// <summary>
/// Dependency-free PNG encoder for RGBA raster images.
/// </summary>
public static class OfficePngWriter {
    private static readonly byte[] PngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    /// <summary>
    /// Encodes an RGBA image as PNG bytes.
    /// </summary>
    public static byte[] Encode(OfficeRasterImage image, OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (image == null) {
            throw new ArgumentNullException(nameof(image));
        }

        return EncodeRgba(image.Width, image.Height, image.GetPixels(), compression);
    }

    /// <summary>Encodes an RGBA image with explicit compression and physical-resolution metadata.</summary>
    public static byte[] Encode(OfficeRasterImage image, OfficePngEncodeOptions options) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
        return EncodeRgba(image.Width, image.Height, image.GetPixels(), options);
    }

    /// <summary>
    /// Encodes raw RGBA pixels as PNG bytes.
    /// </summary>
    public static byte[] EncodeRgba(int width, int height, byte[] rgba, OfficePngCompression compression = OfficePngCompression.Optimal) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height));
        }

        if (rgba == null) {
            throw new ArgumentNullException(nameof(rgba));
        }

        if (rgba.Length != checked(width * height * 4)) {
            throw new ArgumentException("RGBA buffer length does not match image dimensions.", nameof(rgba));
        }

        byte[] scanlines = new byte[height * (1 + width * 4)];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++) {
            scanlines[target++] = 0;
            Buffer.BlockCopy(rgba, source, scanlines, target, width * 4);
            source += width * 4;
            target += width * 4;
        }

        return EncodeScanlines(width, height, 8, 6, scanlines, compression);
    }

    /// <summary>Encodes raw RGBA pixels with explicit compression and physical-resolution metadata.</summary>
    public static byte[] EncodeRgba(int width, int height, byte[] rgba, OfficePngEncodeOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
        ValidateRgba(width, height, rgba);

        byte[] scanlines = CreateRgbaScanlines(width, height, rgba);
        byte[] compressed = options.Compression switch {
            OfficePngCompression.Optimal => DeflateZlib(scanlines),
            OfficePngCompression.Stored => DeflateZlibStored(scanlines),
            _ => throw new ArgumentOutOfRangeException(nameof(options.Compression))
        };
        return CreateFromCompressedScanlines(
            width,
            height,
            8,
            6,
            compressed,
            options.DpiX,
            options.DpiY);
    }

    /// <summary>
    /// Encodes already-filtered PNG scanlines into a PNG file.
    /// </summary>
    /// <remarks>
    /// Each scanline must include its leading PNG filter byte. This entry point is intended for
    /// document adapters that already own source pixel/filter semantics but should still share
    /// OfficeIMO's PNG container, chunk, CRC, and zlib writing.
    /// </remarks>
    public static byte[] EncodeScanlines(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[] scanlines,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        ValidatePngHeader(width, height, bitDepth, colorType);
        if (scanlines == null) {
            throw new ArgumentNullException(nameof(scanlines));
        }

        ValidateScanlineBufferLength(width, height, bitDepth, colorType, scanlines);

        byte[] compressed = compression switch {
            OfficePngCompression.Optimal => DeflateZlib(scanlines),
            OfficePngCompression.Stored => DeflateZlibStored(scanlines),
            _ => throw new ArgumentOutOfRangeException(nameof(compression))
        };
        return CreateFromCompressedScanlines(width, height, bitDepth, colorType, compressed);
    }

    /// <summary>
    /// Wraps zlib-compressed, already-filtered PNG scanlines in a PNG file.
    /// </summary>
    /// <remarks>
    /// The compressed payload is written as the PNG IDAT chunk without decoding or recompressing it.
    /// This is useful when a source document already stores PNG-compatible Flate data.
    /// </remarks>
    public static byte[] CreateFromCompressedScanlines(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[] compressedScanlines) =>
        CreateFromCompressedScanlines(width, height, bitDepth, colorType, compressedScanlines, null, null);

    private static byte[] CreateFromCompressedScanlines(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[] compressedScanlines,
        double? dpiX,
        double? dpiY) {
        ValidatePngHeader(width, height, bitDepth, colorType);
        if (compressedScanlines == null) {
            throw new ArgumentNullException(nameof(compressedScanlines));
        }

        using MemoryStream stream = new MemoryStream();
        stream.Write(PngSignature, 0, PngSignature.Length);
        WriteChunk(stream, "IHDR", BuildIhdr(width, height, bitDepth, colorType));
        if (dpiX.HasValue && dpiY.HasValue) WriteChunk(stream, "pHYs", BuildPhysicalResolution(dpiX.Value, dpiY.Value));
        WriteChunk(stream, "IDAT", compressedScanlines);
        WriteChunk(stream, "IEND", Array.Empty<byte>());
        return stream.ToArray();
    }

    private static void ValidateRgba(int width, int height, byte[] rgba) {
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        if (rgba == null) throw new ArgumentNullException(nameof(rgba));
        if (rgba.Length != checked(width * height * 4)) {
            throw new ArgumentException("RGBA buffer length does not match image dimensions.", nameof(rgba));
        }
    }

    private static byte[] CreateRgbaScanlines(int width, int height, byte[] rgba) {
        byte[] scanlines = new byte[height * (1 + width * 4)];
        int source = 0;
        int target = 0;
        for (int y = 0; y < height; y++) {
            scanlines[target++] = 0;
            Buffer.BlockCopy(rgba, source, scanlines, target, width * 4);
            source += width * 4;
            target += width * 4;
        }
        return scanlines;
    }

    private static byte[] BuildPhysicalResolution(double dpiX, double dpiY) {
        ValidateDpi(dpiX, nameof(dpiX));
        ValidateDpi(dpiY, nameof(dpiY));
        uint pixelsPerMeterX = checked((uint)Math.Round(dpiX / 0.0254D, MidpointRounding.AwayFromZero));
        uint pixelsPerMeterY = checked((uint)Math.Round(dpiY / 0.0254D, MidpointRounding.AwayFromZero));
        byte[] data = new byte[9];
        WriteBigEndianInt32(data, 0, unchecked((int)pixelsPerMeterX));
        WriteBigEndianInt32(data, 4, unchecked((int)pixelsPerMeterY));
        data[8] = 1;
        return data;
    }

    private static void ValidateDpi(double dpi, string paramName) {
        if (dpi < OfficeRasterImageEncoder.PngMinimumDpi ||
            double.IsNaN(dpi) ||
            double.IsInfinity(dpi) ||
            dpi > uint.MaxValue * 0.0254D) {
            throw new ArgumentOutOfRangeException(
                paramName,
                "PNG DPI must be finite and between 0.0127 and " +
                (uint.MaxValue * 0.0254D).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                ".");
        }
    }

    private static void ValidatePngHeader(int width, int height, int bitDepth, int colorType) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height));
        }

        if (bitDepth is not (1 or 2 or 4 or 8 or 16)) {
            throw new ArgumentOutOfRangeException(nameof(bitDepth), "Bit depth must be a valid PNG bit depth.");
        }

        if (colorType is not (0 or 2 or 3 or 4 or 6)) {
            throw new ArgumentOutOfRangeException(nameof(colorType), "Color type must be a valid PNG color type.");
        }

        if (colorType == 3) {
            throw new ArgumentOutOfRangeException(nameof(colorType), "Indexed-color PNGs require a palette (PLTE), which this writer does not emit.");
        }

        bool validForColorType = colorType == 0
            || bitDepth is 8 or 16;
        if (!validForColorType) {
            throw new ArgumentOutOfRangeException(nameof(bitDepth), "Bit depth is not valid for the PNG color type.");
        }
    }

    private static void ValidateScanlineBufferLength(int width, int height, int bitDepth, int colorType, byte[] scanlines) {
        int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
        long rowPayloadBytes = ((long)width * bitsPerPixel + 7L) / 8L;
        long expectedLength = (rowPayloadBytes + 1L) * height;
        if (scanlines.LongLength != expectedLength) {
            throw new ArgumentException(
                "PNG scanline buffer length must match the IHDR width, height, bit depth, and color type including one filter byte per row.",
                nameof(scanlines));
        }
    }

    private static int GetBitsPerPixel(int colorType, int bitDepth) {
        switch (colorType) {
            case 0:
                return bitDepth;
            case 2:
                return bitDepth * 3;
            case 4:
                return bitDepth * 2;
            case 6:
                return bitDepth * 4;
            default:
                throw new ArgumentOutOfRangeException(nameof(colorType), "Color type must be a supported PNG color type.");
        }
    }

    private static byte[] BuildIhdr(int width, int height, int bitDepth, int colorType) {
        byte[] ihdr = new byte[13];
        WriteBigEndianInt32(ihdr, 0, width);
        WriteBigEndianInt32(ihdr, 4, height);
        ihdr[8] = (byte)bitDepth;
        ihdr[9] = (byte)colorType;
        return ihdr;
    }

    private static byte[] DeflateZlib(byte[] data) {
        using MemoryStream stream = new MemoryStream();
        stream.WriteByte(0x78);
        stream.WriteByte(0x9C);
        using (DeflateStream deflate = new DeflateStream(stream, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        stream.WriteByte((byte)((adler >> 24) & 0xFF));
        stream.WriteByte((byte)((adler >> 16) & 0xFF));
        stream.WriteByte((byte)((adler >> 8) & 0xFF));
        stream.WriteByte((byte)(adler & 0xFF));
        return stream.ToArray();
    }

    private static byte[] DeflateZlibStored(byte[] data) {
        using MemoryStream stream = new MemoryStream();
        stream.WriteByte(0x78);
        stream.WriteByte(0x01);

        int offset = 0;
        do {
            int blockLength = Math.Min(65535, data.Length - offset);
            bool final = offset + blockLength >= data.Length;
            stream.WriteByte(final ? (byte)1 : (byte)0);
            stream.WriteByte((byte)(blockLength & 0xFF));
            stream.WriteByte((byte)((blockLength >> 8) & 0xFF));
            ushort nlen = (ushort)~blockLength;
            stream.WriteByte((byte)(nlen & 0xFF));
            stream.WriteByte((byte)((nlen >> 8) & 0xFF));
            stream.Write(data, offset, blockLength);
            offset += blockLength;
        } while (offset < data.Length);

        uint adler = Adler32(data);
        stream.WriteByte((byte)((adler >> 24) & 0xFF));
        stream.WriteByte((byte)((adler >> 16) & 0xFF));
        stream.WriteByte((byte)((adler >> 8) & 0xFF));
        stream.WriteByte((byte)(adler & 0xFF));
        return stream.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static void WriteChunk(Stream stream, string type, byte[] data) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        byte[] length = new byte[4];
        WriteBigEndianInt32(length, 0, data.Length);
        stream.Write(length, 0, length.Length);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        uint crc = Crc32(typeBytes, data);
        byte[] crcBytes = new byte[4];
        WriteBigEndianInt32(crcBytes, 0, unchecked((int)crc));
        stream.Write(crcBytes, 0, crcBytes.Length);
    }

    private static uint Crc32(byte[] type, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < type.Length; i++) {
            crc = UpdateCrc(crc, type[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc(uint crc, byte value) {
        crc ^= value;
        for (int i = 0; i < 8; i++) {
            crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
        }

        return crc;
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)((value >> 24) & 0xFF);
        bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
        bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 3] = (byte)(value & 0xFF);
    }
}
