using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Defines the zlib compression strategy used when encoding already-filtered PNG scanlines.
/// </summary>
public enum OfficePngCompression {
    /// <summary>
    /// Select exact grayscale, RGB, RGBA, or indexed samples and compress
    /// according to the selected encoding effort. All RGBA values,
    /// including hidden colors under transparent pixels, are preserved.
    /// </summary>
    Optimal,

    /// <summary>
    /// Store scanlines in zlib blocks without deflate compression.
    /// </summary>
    Stored,

    /// <summary>Compress scanlines once using the platform's fastest deflate strategy.</summary>
    Fastest
}

/// <summary>
/// Dependency-free PNG encoder for RGBA raster images.
/// </summary>
public static partial class OfficePngWriter {
    private static readonly byte[] PngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    /// <summary>
    /// Encodes an RGBA image as PNG bytes.
    /// </summary>
    public static byte[] Encode(OfficeRasterImage image, OfficePngCompression compression = OfficePngCompression.Optimal) {
        ArgumentNullException.ThrowIfNull(image);

        return EncodeRgba(image.Width, image.Height, image.PixelBuffer, compression);
    }

    /// <summary>Encodes an RGBA image as PNG bytes with cooperative cancellation between scanlines.</summary>
    public static byte[] Encode(
        OfficeRasterImage image,
        System.Threading.CancellationToken cancellationToken,
        OfficePngCompression compression = OfficePngCompression.Optimal) {
        ArgumentNullException.ThrowIfNull(image);
        return EncodeRgbaMaterialized(image.Width, image.Height, image.PixelBuffer,
            compression, null, null, cancellationToken);
    }

    /// <summary>Encodes an RGBA image with explicit compression and physical-resolution metadata.</summary>
    public static byte[] Encode(OfficeRasterImage image, OfficePngEncodeOptions options) {
        ArgumentNullException.ThrowIfNull(image);
        ArgumentNullException.ThrowIfNull(options);
        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
        return EncodeRgba(image.Width, image.Height, image.PixelBuffer, options);
    }

    /// <summary>
    /// Encodes raw RGBA pixels as PNG bytes.
    /// </summary>
    public static byte[] EncodeRgba(int width, int height, byte[] rgba, OfficePngCompression compression = OfficePngCompression.Optimal) {
        return EncodeRgbaMaterialized(
            width, height, rgba, compression, null, null,
            System.Threading.CancellationToken.None);
    }

    /// <summary>Encodes raw RGBA pixels with explicit compression and physical-resolution metadata.</summary>
    public static byte[] EncodeRgba(int width, int height, byte[] rgba, OfficePngEncodeOptions options) {
        ArgumentNullException.ThrowIfNull(options);
        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
        ValidateRgba(width, height, rgba);

        if (options.Compression != OfficePngCompression.Optimal && options.Compression != OfficePngCompression.Stored && options.Compression != OfficePngCompression.Fastest) {
            throw new ArgumentOutOfRangeException(nameof(options.Compression));
        }

        return EncodeRgbaMaterialized(
            width, height, rgba, options.Compression,
            options.WritePhysicalResolution ? options.DpiX : (double?)null,
            options.WritePhysicalResolution ? options.DpiY : (double?)null,
            System.Threading.CancellationToken.None, options.RequireRgbColorSpace, options.RequireGrayColorSpace, options.Effort);
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
        ArgumentNullException.ThrowIfNull(scanlines);

        ValidateScanlineBufferLength(width, height, bitDepth, colorType, scanlines);
        return AssembleScanlines(width, height, bitDepth, colorType, palette: null, scanlines, compression);
    }

    /// <summary>
    /// Encodes already-filtered indexed-color scanlines with an RGB palette (PNG color type 3).
    /// </summary>
    /// <param name="width">Image width in pixels.</param>
    /// <param name="height">Image height in pixels.</param>
    /// <param name="bitDepth">Index bit depth: 1, 2, 4, or 8.</param>
    /// <param name="rgbPalette">RGB triplets, at most 2^bitDepth entries; every index in the scanlines must address one.</param>
    /// <param name="scanlines">Packed indices, each row preceded by its PNG filter byte.</param>
    /// <param name="compression">zlib strategy for the IDAT payload.</param>
    internal static byte[] EncodeIndexedScanlines(
        int width,
        int height,
        int bitDepth,
        byte[] rgbPalette,
        byte[] scanlines,
        OfficePngCompression compression) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        if (bitDepth is not (1 or 2 or 4 or 8)) {
            throw new ArgumentOutOfRangeException(nameof(bitDepth), "Indexed-color PNG bit depth must be 1, 2, 4, or 8.");
        }
        ArgumentNullException.ThrowIfNull(rgbPalette);
        ArgumentNullException.ThrowIfNull(scanlines);
        if (rgbPalette.Length == 0 || rgbPalette.Length % 3 != 0 || rgbPalette.Length / 3 > 1 << bitDepth) {
            throw new ArgumentException("The palette must hold between 1 and 2^bitDepth RGB entries.", nameof(rgbPalette));
        }
        long rowPayloadBytes = ((long)width * bitDepth + 7L) / 8L;
        if (scanlines.LongLength != (rowPayloadBytes + 1L) * height) {
            throw new ArgumentException(
                "PNG scanline buffer length must match the IHDR width, height, and bit depth including one filter byte per row.",
                nameof(scanlines));
        }
        return AssembleScanlines(width, height, bitDepth, 3, rgbPalette, scanlines, compression);
    }

    /// <summary>
    /// Encodes a stored (uncompressed) PNG whose filtered scanlines the caller writes straight into the
    /// output, so no separate scanline buffer is needed. The bytes match the scanline-array overloads with
    /// <see cref="OfficePngCompression.Stored"/>.
    /// </summary>
    /// <param name="width">Image width in pixels.</param>
    /// <param name="height">Image height in pixels.</param>
    /// <param name="bitDepth">Sample bit depth.</param>
    /// <param name="colorType">PNG color type; 3 requires <paramref name="rgbPalette"/>.</param>
    /// <param name="rgbPalette">RGB palette for indexed color, otherwise null.</param>
    /// <param name="writeScanlines">
    /// Writes every row, each preceded by its filter byte, into the zero-filled buffer starting at the given offset.
    /// </param>
    internal static byte[] EncodeStoredScanlines(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[]? rgbPalette,
        Action<byte[], int> writeScanlines) {
        long rowPayloadBytes;
        if (colorType == 3) {
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
            if (bitDepth is not (1 or 2 or 4 or 8)) {
                throw new ArgumentOutOfRangeException(nameof(bitDepth), "Indexed-color PNG bit depth must be 1, 2, 4, or 8.");
            }
            ArgumentNullException.ThrowIfNull(rgbPalette);
            if (rgbPalette.Length == 0 || rgbPalette.Length % 3 != 0 || rgbPalette.Length / 3 > 1 << bitDepth) {
                throw new ArgumentException("The palette must hold between 1 and 2^bitDepth RGB entries.", nameof(rgbPalette));
            }
            rowPayloadBytes = ((long)width * bitDepth + 7L) / 8L;
        } else {
            ValidatePngHeader(width, height, bitDepth, colorType);
            if (rgbPalette is not null) throw new ArgumentException("Only indexed-color PNGs carry a palette.", nameof(rgbPalette));
            rowPayloadBytes = ((long)width * GetBitsPerPixel(colorType, bitDepth) + 7L) / 8L;
        }
        ArgumentNullException.ThrowIfNull(writeScanlines);
        int scanlineLength = checked((int)((rowPayloadBytes + 1L) * height));
        return AssemblePng(width, height, bitDepth, colorType, rgbPalette, compressed: null, storedScanlines: null,
            scanlineLength, writeScanlines);
    }

    private static byte[] AssembleScanlines(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[]? palette,
        byte[] scanlines,
        OfficePngCompression compression) {
        if (compression is not (OfficePngCompression.Optimal or OfficePngCompression.Fastest or OfficePngCompression.Stored)) {
            throw new ArgumentOutOfRangeException(nameof(compression));
        }
        return compression == OfficePngCompression.Stored
            ? AssemblePng(width, height, bitDepth, colorType, palette, compressed: null, storedScanlines: scanlines)
            : AssemblePng(width, height, bitDepth, colorType, palette,
                DeflateZlib(scanlines, compression == OfficePngCompression.Fastest ? CompressionLevel.Fastest : CompressionLevel.Optimal),
                storedScanlines: null);
    }

    // Writes the whole PNG into one exactly sized array. Stored zlib blocks are produced in place,
    // so a large uncompressed image costs one output allocation instead of a chain of growing copies.
    private static byte[] AssemblePng(
        int width,
        int height,
        int bitDepth,
        int colorType,
        byte[]? palette,
        byte[]? compressed,
        byte[]? storedScanlines,
        int writtenScanlineLength = 0,
        Action<byte[], int>? writeScanlines = null) {
        byte[] scanlines = storedScanlines ?? Array.Empty<byte>();
        int scanlineLength = writeScanlines is null ? scanlines.Length : writtenScanlineLength;
        int storedBlocks = Math.Max(1, (scanlineLength + 65534) / 65535);
        int idatLength = compressed?.Length ?? checked(2 + storedBlocks * 5 + scanlineLength + 4);
        int paletteChunk = palette is null ? 0 : checked(12 + palette.Length);
        byte[] png = new byte[checked(PngSignature.Length + 25 + paletteChunk + 12 + idatLength + 12)];
        Buffer.BlockCopy(PngSignature, 0, png, 0, PngSignature.Length);
        int position = PngSignature.Length;
        int data = BeginChunk(png, ref position, "IHDR", 13);
        WriteBigEndianInt32(png, data, width);
        WriteBigEndianInt32(png, data + 4, height);
        png[data + 8] = (byte)bitDepth;
        png[data + 9] = (byte)colorType;
        position = data + 13;
        CompleteChunk(png, ref position, data, 13);
        if (palette is not null) {
            data = BeginChunk(png, ref position, "PLTE", palette.Length);
            Buffer.BlockCopy(palette, 0, png, data, palette.Length);
            position = data + palette.Length;
            CompleteChunk(png, ref position, data, palette.Length);
        }
        data = BeginChunk(png, ref position, "IDAT", idatLength);
        if (compressed is not null) {
            Buffer.BlockCopy(compressed, 0, png, data, compressed.Length);
            position = data + compressed.Length;
        } else if (writeScanlines is not null) {
            position = data;
            png[position++] = 0x78;
            png[position++] = 0x01;
            // The caller writes the scanlines contiguously after room for every block header; each block then
            // moves left into place. A block never overwrites a later block's source bytes before it moves.
            int contiguous = position + storedBlocks * 5;
            writeScanlines(png, contiguous);
            uint a = 1;
            uint b = 0;
            UpdateAdler32(png, contiguous, scanlineLength, ref a, ref b);
            int offset = 0;
            do {
                int blockLength = Math.Min(65535, scanlineLength - offset);
                png[position++] = offset + blockLength >= scanlineLength ? (byte)1 : (byte)0;
                png[position++] = (byte)(blockLength & 0xFF);
                png[position++] = (byte)((blockLength >> 8) & 0xFF);
                ushort nlen = (ushort)~blockLength;
                png[position++] = (byte)(nlen & 0xFF);
                png[position++] = (byte)((nlen >> 8) & 0xFF);
                if (position != contiguous + offset) Buffer.BlockCopy(png, contiguous + offset, png, position, blockLength);
                position += blockLength;
                offset += blockLength;
            } while (offset < scanlineLength);
            WriteBigEndianInt32(png, position, unchecked((int)((b << 16) | a)));
            position += 4;
        } else {
            position = data;
            png[position++] = 0x78;
            png[position++] = 0x01;
            int offset = 0;
            do {
                int blockLength = Math.Min(65535, scanlines.Length - offset);
                png[position++] = offset + blockLength >= scanlines.Length ? (byte)1 : (byte)0;
                png[position++] = (byte)(blockLength & 0xFF);
                png[position++] = (byte)((blockLength >> 8) & 0xFF);
                ushort nlen = (ushort)~blockLength;
                png[position++] = (byte)(nlen & 0xFF);
                png[position++] = (byte)((nlen >> 8) & 0xFF);
                Buffer.BlockCopy(scanlines, offset, png, position, blockLength);
                position += blockLength;
                offset += blockLength;
            } while (offset < scanlines.Length);
            WriteBigEndianInt32(png, position, unchecked((int)Adler32(scanlines)));
            position += 4;
        }
        CompleteChunk(png, ref position, data, idatLength);
        data = BeginChunk(png, ref position, "IEND", 0);
        CompleteChunk(png, ref position, data, 0);
        return png;
    }

    private static int BeginChunk(byte[] png, ref int position, string type, int length) {
        WriteBigEndianInt32(png, position, length);
        for (int i = 0; i < 4; i++) png[position + 4 + i] = (byte)type[i];
        position += 8;
        return position;
    }

    private static void CompleteChunk(byte[] png, ref int position, int dataOffset, int length) {
        uint crc = OfficePngCrc32.Compute(png, dataOffset - 4, length + 4);
        WriteBigEndianInt32(png, position, unchecked((int)crc));
        position += 4;
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
        byte[] compressedScanlines) {
        ValidatePngHeader(width, height, bitDepth, colorType);
        ArgumentNullException.ThrowIfNull(compressedScanlines);
        return AssemblePng(width, height, bitDepth, colorType, palette: null, compressedScanlines, storedScanlines: null);
    }

    private static void ValidateRgba(int width, int height, byte[] rgba) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);
        ArgumentNullException.ThrowIfNull(rgba);
        if (rgba.Length != checked(width * height * 4)) {
            throw new ArgumentException("RGBA buffer length does not match image dimensions.", nameof(rgba));
        }
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
            dpi > OfficeRasterImageEncoder.PngMaximumDpi) {
            throw new ArgumentOutOfRangeException(
                paramName,
                "PNG DPI must be finite and between 0.0127 and " +
                OfficeRasterImageEncoder.PngMaximumDpi.ToString(System.Globalization.CultureInfo.InvariantCulture) +
                ".");
        }
    }

    private static void ValidatePngHeader(int width, int height, int bitDepth, int colorType) {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(width);

        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(height);

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

    private static byte[] DeflateZlib(byte[] data, CompressionLevel level = CompressionLevel.Optimal) {
        using MemoryStream stream = new MemoryStream();
        stream.WriteByte(0x78);
        stream.WriteByte(0x9C);
        using (DeflateStream deflate = new DeflateStream(stream, level, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }

        uint adler = Adler32(data);
        stream.WriteByte((byte)((adler >> 24) & 0xFF));
        stream.WriteByte((byte)((adler >> 16) & 0xFF));
        stream.WriteByte((byte)((adler >> 8) & 0xFF));
        stream.WriteByte((byte)(adler & 0xFF));
        return stream.ToArray();
    }

    private static uint Adler32(byte[] data) {
        uint a = 1;
        uint b = 0;
        UpdateAdler32(data, 0, data.Length, ref a, ref b);
        return (b << 16) | a;
    }

    private static void UpdateAdler32(
        byte[] data,
        int offset,
        int count,
        ref uint a,
        ref uint b,
        System.Threading.CancellationToken cancellationToken = default,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver = null) {
        const int maximumChunk = 5552;
        int remaining = count;
        while (remaining > 0) {
            checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
            cancellationToken.ThrowIfCancellationRequested();
            int chunk = Math.Min(maximumChunk, remaining);
            OfficeIMO.Core.Internal.OfficeZlibCodec.AppendAdler32Block(data, offset, chunk, ref a, ref b);
            offset += chunk;
            remaining -= chunk;
        }
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
        uint crc = OfficePngCrc32.Append(OfficePngCrc32.Begin(), type, 0, type.Length);
        crc = OfficePngCrc32.Append(crc, data, 0, data.Length);
        return OfficePngCrc32.Complete(crc);
    }

    private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)((value >> 24) & 0xFF);
        bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
        bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
        bytes[offset + 3] = (byte)(value & 0xFF);
    }
}
