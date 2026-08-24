using System;
using System.IO;
using System.Text;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free PNG decoder for supported PNG images.
/// </summary>
public static class OfficePngReader {
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    /// <summary>
    /// Inspects a PNG container without decoding pixels and reports its declared APNG frame count.
    /// Static PNG files report one frame.
    /// </summary>
    public static bool TryGetFrameCount(byte[]? bytes, out int frameCount) {
        return TryGetFrameCount(bytes, CancellationToken.None, out frameCount);
    }

    internal static bool TryGetFrameCount(
        byte[]? bytes,
        CancellationToken cancellationToken,
        out int frameCount) {
        frameCount = 0;
        try {
            return OfficePngContainerValidator.TryValidate(bytes, cancellationToken, out frameCount, out _);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            frameCount = 0;
            return false;
        }
    }

    /// <summary>
    /// Attempts to decode a PNG image into an RGBA raster buffer.
    /// </summary>
    public static bool TryDecode(byte[] bytes, out OfficeRasterImage? image) {
        return TryDecode(bytes, CancellationToken.None, out image);
    }

    internal static bool TryDecode(
        byte[] bytes,
        CancellationToken cancellationToken,
        out OfficeRasterImage? image) {
        image = null;
        try {
            if (!TryReadPayload(bytes, cancellationToken, includeRgbaOutput: true, out PngPayload payload)) {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            OfficeRasterImage result = new OfficeRasterImage(payload.Width, payload.Height);
            if (payload.InterlaceMethod == 0) {
                DecodeScanlines(payload, result, cancellationToken);
            } else if (!DecodeAdam7Scanlines(payload, result, cancellationToken)) {
                return false;
            }

            image = result;
            return true;
        } catch (OperationCanceledException) {
            throw;
        } catch {
            image = null;
            return false;
        }
    }

    private static void DecodeScanlines(
        PngPayload payload,
        OfficeRasterImage result,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] previous = new byte[payload.Stride];
        cancellationToken.ThrowIfCancellationRequested();
        byte[] current = new byte[payload.Stride];
        int sourceOffset = 0;
        for (int y = 0; y < payload.Height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            int filter = payload.Scanlines[sourceOffset++];
            CopyBytes(payload.Scanlines, sourceOffset, current, 0, payload.Stride, cancellationToken);
            sourceOffset += payload.Stride;
            Unfilter(current, previous, payload.BytesPerPixel, filter, cancellationToken);
            ExpandScanline(current, payload.Width, y, payload.ColorType, payload.BitDepth,
                payload.Palette, payload.Transparency, result, cancellationToken);

            byte[] temp = previous;
            previous = current;
            current = temp;
            ClearBytes(current, cancellationToken);
        }
    }

    private static bool DecodeAdam7Scanlines(
        PngPayload payload,
        OfficeRasterImage result,
        CancellationToken cancellationToken) {
        int[] startX = { 0, 4, 0, 2, 0, 1, 0 };
        int[] startY = { 0, 0, 4, 0, 2, 0, 1 };
        int[] stepX = { 8, 8, 4, 4, 2, 2, 1 };
        int[] stepY = { 8, 8, 8, 4, 4, 2, 2 };
        int sourceOffset = 0;
        int bitsPerPixel = GetBitsPerPixel(payload.ColorType, payload.BitDepth);
        for (int pass = 0; pass < 7; pass++) {
            int passWidth = GetAdam7PassLength(payload.Width, startX[pass], stepX[pass]);
            int passHeight = GetAdam7PassLength(payload.Height, startY[pass], stepY[pass]);
            if (passWidth == 0 || passHeight == 0) continue;
            int stride = OfficeRasterGuards.EnsureByteCount(
                (((long)passWidth * bitsPerPixel) + 7L) / 8L,
                "PNG Adam7 scanline dimensions exceed size limits.");
            cancellationToken.ThrowIfCancellationRequested();
            byte[] previous = new byte[stride];
            cancellationToken.ThrowIfCancellationRequested();
            byte[] current = new byte[stride];
            for (int passY = 0; passY < passHeight; passY++) {
                if ((passY & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (sourceOffset > payload.Scanlines.Length - stride - 1) return false;
                int filter = payload.Scanlines[sourceOffset++];
                CopyBytes(payload.Scanlines, sourceOffset, current, 0, stride, cancellationToken);
                sourceOffset += stride;
                Unfilter(current, previous, payload.BytesPerPixel, filter, cancellationToken);
                ExpandScanline(current, passWidth, startY[pass] + passY * stepY[pass],
                    payload.ColorType, payload.BitDepth, payload.Palette, payload.Transparency,
                    result, cancellationToken, startX[pass], stepX[pass]);

                byte[] temp = previous;
                previous = current;
                current = temp;
                ClearBytes(current, cancellationToken);
            }
        }
        return sourceOffset == payload.Scanlines.Length;
    }

    /// <summary>Validates that a PNG has a complete, bounded scanline payload without allocating an RGBA raster.</summary>
    internal static bool TryValidateDecodedPayload(byte[] bytes) =>
        TryValidateDecodedPayload(bytes, CancellationToken.None);

    internal static bool TryValidateDecodedPayload(byte[] bytes, CancellationToken cancellationToken) {
        try {
            if (!TryReadPayload(bytes, cancellationToken, includeRgbaOutput: false, out PngPayload payload)) return false;
            if (!ValidatePayloadScanlines(payload, cancellationToken)) return false;
            long retainedPayloadBytes = checked(
                payload.Scanlines.LongLength +
                (payload.Palette?.LongLength ?? 0L) +
                (payload.Transparency?.LongLength ?? 0L));
            return OfficePngAnimationValidator.TryValidateAdditionalFrames(
                bytes, retainedPayloadBytes, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        }
    }

    internal static bool TryValidateCompressedPayload(
        byte[] compressed,
        int width,
        int height,
        int bitDepth,
        int colorType,
        int interlaceMethod,
        byte[]? palette,
        CancellationToken cancellationToken = default) {
        try {
            if (compressed == null || compressed.Length < 6 ||
                !IsSupportedColorLayout(colorType, bitDepth, palette) ||
                !OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
                return false;
            }

            int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
            int stride = OfficeRasterGuards.EnsureByteCount(
                (((long)width * bitsPerPixel) + 7L) / 8L,
                "PNG scanline dimensions exceed size limits.");
            int expectedScanlineBytes = interlaceMethod == 0
                ? OfficeRasterGuards.EnsureByteCount((long)(stride + 1) * height, "PNG decompressed data exceeds size limits.")
                : GetExpectedAdam7ScanlineBytes(width, height, bitsPerPixel);
            if (!IsDecodeWorkingSetWithinLimit(
                    compressed.LongLength,
                    compressedBufferBytes: 0L,
                    compressedCopyBytes: 0L,
                    width,
                    height,
                    stride,
                    expectedScanlineBytes,
                    palette?.LongLength ?? 0L,
                    transparencyBytes: 0L,
                    includeRgbaOutput: false)) return false;
            byte[] scanlines = OfficeZlibCodec.Decompress(
                compressed, expectedScanlineBytes, expectedScanlineBytes, cancellationToken);
            var payload = new PngPayload(
                width,
                height,
                bitDepth,
                colorType,
                interlaceMethod,
                Math.Max(1, (bitsPerPixel + 7) / 8),
                stride,
                palette,
                transparency: null,
                scanlines);
            return ValidatePayloadScanlines(payload, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            return false;
        }
    }

    private static bool ValidatePayloadScanlines(
        PngPayload payload,
        CancellationToken cancellationToken = default) {
        if (payload.InterlaceMethod == 1) return ValidateAdam7Scanlines(payload, cancellationToken);
        return ValidateScanlines(
                   payload, payload.Width, payload.Height, payload.Stride, 0, cancellationToken, out int consumed) &&
               consumed == payload.Scanlines.Length;
    }

    internal static bool TryGetValidationWorkingSetBytes(
        int width,
        int height,
        int bitDepth,
        int colorType,
        int interlaceMethod,
        byte[]? palette,
        out long workingSetBytes) {
        workingSetBytes = 0L;
        try {
            if (!IsSupportedColorLayout(colorType, bitDepth, palette) ||
                !OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) return false;
            int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
            long stride = (((long)width * bitsPerPixel) + 7L) / 8L;
            int expectedScanlineBytes = interlaceMethod == 0
                ? OfficeRasterGuards.EnsureByteCount((stride + 1L) * height, "PNG decompressed data exceeds size limits.")
                : GetExpectedAdam7ScanlineBytes(width, height, bitsPerPixel);
            workingSetBytes = checked(expectedScanlineBytes + stride * 2L);
            return workingSetBytes <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (Exception exception) when (
            exception is ArgumentException || exception is FormatException || exception is OverflowException) {
            return false;
        }
    }

    private static bool ValidateAdam7Scanlines(PngPayload payload, CancellationToken cancellationToken) {
        int[] startX = { 0, 4, 0, 2, 0, 1, 0 };
        int[] startY = { 0, 0, 4, 0, 2, 0, 1 };
        int[] stepX = { 8, 8, 4, 4, 2, 2, 1 };
        int[] stepY = { 8, 8, 8, 4, 4, 2, 2 };
        int sourceOffset = 0;
        int bitsPerPixel = GetBitsPerPixel(payload.ColorType, payload.BitDepth);
        for (int pass = 0; pass < 7; pass++) {
            int passWidth = GetAdam7PassLength(payload.Width, startX[pass], stepX[pass]);
            int passHeight = GetAdam7PassLength(payload.Height, startY[pass], stepY[pass]);
            if (passWidth == 0 || passHeight == 0) continue;
            int stride = OfficeRasterGuards.EnsureByteCount(
                (((long)passWidth * bitsPerPixel) + 7L) / 8L,
                "PNG Adam7 scanline dimensions exceed size limits.");
            if (!ValidateScanlines(
                    payload, passWidth, passHeight, stride, sourceOffset, cancellationToken, out int consumed)) return false;
            sourceOffset = consumed;
        }
        return sourceOffset == payload.Scanlines.Length;
    }

    private static bool ValidateScanlines(
        PngPayload payload,
        int width,
        int height,
        int stride,
        int sourceOffset,
        CancellationToken cancellationToken,
        out int consumed) {
        consumed = sourceOffset;
        if ((long)sourceOffset + ((long)stride + 1L) * height > payload.Scanlines.Length) return false;
        cancellationToken.ThrowIfCancellationRequested();
        byte[] previous = new byte[stride];
        cancellationToken.ThrowIfCancellationRequested();
        byte[] current = new byte[stride];
        for (int y = 0; y < height; y++) {
            if ((y & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            int filter = payload.Scanlines[consumed++];
            CopyBytes(payload.Scanlines, consumed, current, 0, stride, cancellationToken);
            consumed += stride;
            Unfilter(current, previous, payload.BytesPerPixel, filter, cancellationToken);
            if (payload.ColorType == 3) {
                int paletteEntries = payload.Palette!.Length / 3;
                for (int x = 0; x < width; x++) {
                    if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (GetPackedSample(current, x, payload.BitDepth) >= paletteEntries) return false;
                }
            }

            byte[] temp = previous;
            previous = current;
            current = temp;
            ClearBytes(current, cancellationToken);
        }
        return true;
    }

    private static int GetAdam7PassLength(int total, int start, int step) =>
        total <= start ? 0 : (total - start + step - 1) / step;

    private static int GetExpectedAdam7ScanlineBytes(int width, int height, int bitsPerPixel) {
        int[] startX = { 0, 4, 0, 2, 0, 1, 0 };
        int[] startY = { 0, 0, 4, 0, 2, 0, 1 };
        int[] stepX = { 8, 8, 4, 4, 2, 2, 1 };
        int[] stepY = { 8, 8, 8, 4, 4, 2, 2 };
        long expected = 0L;
        for (int pass = 0; pass < 7; pass++) {
            int passWidth = GetAdam7PassLength(width, startX[pass], stepX[pass]);
            int passHeight = GetAdam7PassLength(height, startY[pass], stepY[pass]);
            if (passWidth == 0 || passHeight == 0) continue;
            long stride = (((long)passWidth * bitsPerPixel) + 7L) / 8L;
            expected += (stride + 1L) * passHeight;
        }
        return OfficeRasterGuards.EnsureByteCount(expected, "PNG decompressed Adam7 data exceeds size limits.");
    }

    private static bool TryReadPayload(
        byte[]? bytes,
        CancellationToken cancellationToken,
        bool includeRgbaOutput,
        out PngPayload payload) {
        payload = null!;
        if (bytes == null ||
            !OfficePngContainerValidator.TryValidate(bytes, cancellationToken, out _, out _)) return false;
        cancellationToken.ThrowIfCancellationRequested();

        int width = 0;
        int height = 0;
        int bitDepth = 0;
        int colorType = 0;
        int compressionMethod = 0;
        int filterMethod = 0;
        int interlaceMethod = 0;
        byte[]? palette = null;
        byte[]? transparency = null;
        long compressedLength = 0L;
        int offset = Signature.Length;
        while (offset + 12 <= bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadBigEndianInt32(bytes, offset);
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            int dataOffset = offset + 8;
            if (type == "IHDR") {
                width = ReadBigEndianInt32(bytes, dataOffset);
                height = ReadBigEndianInt32(bytes, dataOffset + 4);
                bitDepth = bytes[dataOffset + 8];
                colorType = bytes[dataOffset + 9];
                compressionMethod = bytes[dataOffset + 10];
                filterMethod = bytes[dataOffset + 11];
                interlaceMethod = bytes[dataOffset + 12];
            } else if (type == "PLTE") {
                palette = new byte[OfficeRasterGuards.EnsureByteCount(length, "PNG palette exceeds size limits.")];
                Buffer.BlockCopy(bytes, dataOffset, palette, 0, length);
            } else if (type == "tRNS") {
                transparency = new byte[OfficeRasterGuards.EnsureByteCount(length, "PNG transparency data exceeds size limits.")];
                Buffer.BlockCopy(bytes, dataOffset, transparency, 0, length);
            } else if (type == "IDAT") {
                compressedLength = checked(compressedLength + length);
            } else if (type == "IEND") {
                break;
            }
            offset += 12 + length;
        }

        if (width <= 0 || height <= 0 || compressionMethod != 0 || filterMethod != 0 || interlaceMethod > 1 ||
            !IsSupportedColorLayout(colorType, bitDepth, palette) ||
            !OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
            return false;
        }

        int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
        int stride = OfficeRasterGuards.EnsureByteCount(
            (((long)width * bitsPerPixel) + 7L) / 8L,
            "PNG scanline dimensions exceed size limits.");
        int expectedScanlineBytes = interlaceMethod == 0
            ? OfficeRasterGuards.EnsureByteCount((long)(stride + 1) * height, "PNG decompressed data exceeds size limits.")
            : GetExpectedAdam7ScanlineBytes(width, height, bitsPerPixel);
        int compressedByteCount = OfficeRasterGuards.EnsureByteCount(
            compressedLength,
            "PNG compressed image data exceeds size limits.");
        if (compressedByteCount < 6) return false;
        if (!IsDecodeWorkingSetWithinLimit(
                bytes.LongLength,
                compressedByteCount,
                compressedCopyBytes: 0L,
                width,
                height,
                stride,
                expectedScanlineBytes,
                palette?.LongLength ?? 0L,
                transparency?.LongLength ?? 0L,
                includeRgbaOutput)) return false;
        var compressed = new byte[compressedByteCount];
        int compressedOffset = 0;
        offset = Signature.Length;
        while (offset + 12 <= bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int length = ReadBigEndianInt32(bytes, offset);
            int dataOffset = offset + 8;
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == "IDAT") {
                CopyBytes(bytes, dataOffset, compressed, compressedOffset, length, cancellationToken);
                compressedOffset += length;
            }
            offset += 12 + length;
            if (type == "IEND") break;
        }
        if (compressedOffset != compressed.Length) return false;
        byte[] scanlines = OfficeZlibCodec.Decompress(
            compressed, expectedScanlineBytes, expectedScanlineBytes, cancellationToken);
        payload = new PngPayload(
            width,
            height,
            bitDepth,
            colorType,
            interlaceMethod,
            Math.Max(1, (bitsPerPixel + 7) / 8),
            stride,
            palette,
            transparency,
            scanlines);
        return true;
    }

    internal static bool IsDecodeWorkingSetWithinLimit(
        long encodedBytes,
        long compressedBufferBytes,
        long compressedCopyBytes,
        int width,
        int height,
        int stride,
        long scanlineBytes,
        long paletteBytes,
        long transparencyBytes,
        bool includeRgbaOutput) {
        if (encodedBytes < 0L || compressedBufferBytes < 0L || compressedCopyBytes < 0L ||
            width < 1 || height < 1 || stride < 1 || scanlineBytes < 0L ||
            paletteBytes < 0L || transparencyBytes < 0L) return false;
        try {
            long metadataBytes = checked(paletteBytes + transparencyBytes + 64L * 1024L);
            long payloadPeakBytes = checked(
                encodedBytes + compressedBufferBytes + compressedCopyBytes + scanlineBytes + metadataBytes);
            long outputBytes = includeRgbaOutput ? checked((long)width * height * 4L) : 0L;
            long decodePeakBytes = checked(
                encodedBytes + scanlineBytes + outputBytes + stride * 2L + metadataBytes);
            return Math.Max(payloadPeakBytes, decodePeakBytes) <= OfficeRasterGuards.MaximumDecodedBytes;
        } catch (OverflowException) {
            return false;
        }
    }

    private sealed class PngPayload {
        internal PngPayload(
            int width,
            int height,
            int bitDepth,
            int colorType,
            int interlaceMethod,
            int bytesPerPixel,
            int stride,
            byte[]? palette,
            byte[]? transparency,
            byte[] scanlines) {
            Width = width;
            Height = height;
            BitDepth = bitDepth;
            ColorType = colorType;
            InterlaceMethod = interlaceMethod;
            BytesPerPixel = bytesPerPixel;
            Stride = stride;
            Palette = palette;
            Transparency = transparency;
            Scanlines = scanlines;
        }

        internal int Width { get; }
        internal int Height { get; }
        internal int BitDepth { get; }
        internal int ColorType { get; }
        internal int InterlaceMethod { get; }
        internal int BytesPerPixel { get; }
        internal int Stride { get; }
        internal byte[]? Palette { get; }
        internal byte[]? Transparency { get; }
        internal byte[] Scanlines { get; }
    }

    private static bool IsSupportedColorLayout(int colorType, int bitDepth, byte[]? palette) {
        switch (colorType) {
            case 0:
                return bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8 || bitDepth == 16;
            case 2:
            case 4:
            case 6:
                return bitDepth == 8 || bitDepth == 16;
            case 3:
                return (bitDepth == 1 || bitDepth == 2 || bitDepth == 4 || bitDepth == 8) &&
                       palette != null &&
                       palette.Length >= 3 &&
                       palette.Length % 3 == 0;
            default:
                return false;
        }
    }

    private static int GetBitsPerPixel(int colorType, int bitDepth) {
        switch (colorType) {
            case 0:
            case 3:
                return bitDepth;
            case 2:
                return bitDepth * 3;
            case 4:
                return bitDepth * 2;
            case 6:
                return bitDepth * 4;
            default:
                throw new InvalidDataException("Unsupported PNG color type.");
        }
    }

    private static void ExpandScanline(
        byte[] current,
        int width,
        int y,
        int colorType,
        int bitDepth,
        byte[]? palette,
        byte[]? transparency,
        OfficeRasterImage image,
        CancellationToken cancellationToken,
        int destinationStartX = 0,
        int destinationStepX = 1) {
        if (colorType == 6 && bitDepth == 8 && destinationStartX == 0 && destinationStepX == 1) {
            CopyBytes(current, 0, image.PixelBuffer, checked(y * width * 4), checked(width * 4),
                cancellationToken);
            return;
        }

        for (int x = 0; x < width; x++) {
            if ((x & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            OfficeColor color;
            switch (colorType) {
                case 0:
                    color = ExpandGrayscale(GetGrayscaleSample(current, x, bitDepth), bitDepth, transparency);
                    break;
                case 2:
                    color = ExpandTrueColor(current, x * (bitDepth == 16 ? 6 : 3), bitDepth, transparency);
                    break;
                case 3:
                    color = ExpandPalette(GetPackedSample(current, x, bitDepth), palette!, transparency);
                    break;
                case 4:
                    color = ExpandGrayscaleAlpha(current, x * (bitDepth == 16 ? 4 : 2), bitDepth);
                    break;
                case 6:
                    color = ExpandTrueColorAlpha(current, x * (bitDepth == 16 ? 8 : 4), bitDepth);
                    break;
                default:
                    throw new InvalidDataException("Unsupported PNG color type.");
            }

            image.SetPixel(destinationStartX + x * destinationStepX, y, color);
        }
    }

    private static OfficeColor ExpandGrayscale(int sample, int bitDepth, byte[]? transparency) {
        byte gray = ScaleSample(sample, bitDepth);
        return OfficeColor.FromRgba(gray, gray, gray, IsTransparentGray(sample, transparency) ? (byte)0 : (byte)255);
    }

    private static OfficeColor ExpandGrayscaleAlpha(byte[] current, int sourcePixel, int bitDepth) {
        int graySample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel) : current[sourcePixel];
        int alphaSample = bitDepth == 16 ? ReadBigEndianUInt16(current, sourcePixel + 2) : current[sourcePixel + 1];
        byte gray = ScaleSample(graySample, bitDepth);
        return OfficeColor.FromRgba(gray, gray, gray, ScaleSample(alphaSample, bitDepth));
    }

    private static OfficeColor ExpandTrueColor(byte[] current, int sourcePixel, int bitDepth, byte[]? transparency) {
        int red;
        int green;
        int blue;
        if (bitDepth == 16) {
            red = ReadBigEndianUInt16(current, sourcePixel);
            green = ReadBigEndianUInt16(current, sourcePixel + 2);
            blue = ReadBigEndianUInt16(current, sourcePixel + 4);
        } else {
            red = current[sourcePixel];
            green = current[sourcePixel + 1];
            blue = current[sourcePixel + 2];
        }

        return OfficeColor.FromRgba(ScaleSample(red, bitDepth), ScaleSample(green, bitDepth), ScaleSample(blue, bitDepth), IsTransparentRgb(red, green, blue, transparency) ? (byte)0 : (byte)255);
    }

    private static OfficeColor ExpandTrueColorAlpha(byte[] current, int sourcePixel, int bitDepth) {
        if (bitDepth == 16) {
            return OfficeColor.FromRgba(
                ScaleSample(ReadBigEndianUInt16(current, sourcePixel), bitDepth),
                ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 2), bitDepth),
                ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 4), bitDepth),
                ScaleSample(ReadBigEndianUInt16(current, sourcePixel + 6), bitDepth));
        }

        return OfficeColor.FromRgba(current[sourcePixel], current[sourcePixel + 1], current[sourcePixel + 2], current[sourcePixel + 3]);
    }

    private static OfficeColor ExpandPalette(int index, byte[] palette, byte[]? transparency) {
        int paletteOffset = index * 3;
        if (paletteOffset + 2 >= palette.Length) {
            throw new InvalidDataException("PNG palette index is outside PLTE.");
        }

        return OfficeColor.FromRgba(palette[paletteOffset], palette[paletteOffset + 1], palette[paletteOffset + 2], transparency != null && index < transparency.Length ? transparency[index] : (byte)255);
    }

    private static int GetPackedSample(byte[] current, int x, int bitDepth) {
        if (bitDepth == 8) return current[x];
        int samplesPerByte = 8 / bitDepth;
        int shift = (samplesPerByte - 1 - (x % samplesPerByte)) * bitDepth;
        int mask = (1 << bitDepth) - 1;
        return (current[x / samplesPerByte] >> shift) & mask;
    }

    private static int GetGrayscaleSample(byte[] current, int x, int bitDepth) =>
        bitDepth == 16 ? ReadBigEndianUInt16(current, x * 2) : bitDepth == 8 ? current[x] : GetPackedSample(current, x, bitDepth);

    private static int ReadBigEndianInt32(byte[] bytes, int offset) =>
        (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];

    private static int ReadBigEndianUInt16(byte[] bytes, int offset) => (bytes[offset] << 8) | bytes[offset + 1];

    private static byte ScaleSample(int sample, int bitDepth) {
        if (bitDepth == 8) return (byte)sample;
        int max = (1 << bitDepth) - 1;
        return (byte)Math.Round(sample * 255D / max);
    }

    private static bool IsTransparentGray(int sample, byte[]? transparency) =>
        transparency != null && transparency.Length >= 2 && sample == ((transparency[0] << 8) | transparency[1]);

    private static bool IsTransparentRgb(int red, int green, int blue, byte[]? transparency) =>
        transparency != null &&
        transparency.Length >= 6 &&
        red == ((transparency[0] << 8) | transparency[1]) &&
        green == ((transparency[2] << 8) | transparency[3]) &&
        blue == ((transparency[4] << 8) | transparency[5]);

    internal static void Unfilter(
        byte[] current,
        byte[] previous,
        int bytesPerPixel,
        int filter,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (filter) {
            case 0:
                return;
            case 1:
                for (int index = bytesPerPixel; index < current.Length; index++) {
                    if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    current[index] = unchecked((byte)(current[index] + current[index - bytesPerPixel]));
                }
                return;
            case 2:
                for (int index = 0; index < current.Length; index++) {
                    if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    current[index] = unchecked((byte)(current[index] + previous[index]));
                }
                return;
            case 3:
                int prefixLength = Math.Min(bytesPerPixel, current.Length);
                for (int index = 0; index < prefixLength; index++) {
                    current[index] = unchecked((byte)(current[index] + (previous[index] / 2)));
                }
                for (int index = bytesPerPixel; index < current.Length; index++) {
                    if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    current[index] = unchecked((byte)(current[index] + ((current[index - bytesPerPixel] + previous[index]) / 2)));
                }
                return;
            case 4:
                int firstPixelBytes = Math.Min(bytesPerPixel, current.Length);
                for (int index = 0; index < firstPixelBytes; index++) {
                    current[index] = unchecked((byte)(current[index] + previous[index]));
                }
                for (int index = bytesPerPixel; index < current.Length; index++) {
                    if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
                    current[index] = unchecked((byte)(current[index] + Paeth(
                        current[index - bytesPerPixel],
                        previous[index],
                        previous[index - bytesPerPixel])));
                }
                return;
            default:
                throw new InvalidDataException("Unsupported PNG filter.");
        }
    }

    internal static void CopyBytes(
        byte[] source,
        int sourceOffset,
        byte[] destination,
        int destinationOffset,
        int count,
        CancellationToken cancellationToken) {
        const int copyChunkBytes = 64 * 1024;
        int remaining = count;
        while (remaining > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            int copy = Math.Min(remaining, copyChunkBytes);
            Buffer.BlockCopy(source, sourceOffset, destination, destinationOffset, copy);
            sourceOffset += copy;
            destinationOffset += copy;
            remaining -= copy;
        }
    }

    internal static void ClearBytes(byte[] bytes, CancellationToken cancellationToken) {
        const int clearChunkBytes = 64 * 1024;
        for (int offset = 0; offset < bytes.Length; offset += clearChunkBytes) {
            cancellationToken.ThrowIfCancellationRequested();
            Array.Clear(bytes, offset, Math.Min(clearChunkBytes, bytes.Length - offset));
        }
    }

    private static int Paeth(int left, int up, int upLeft) {
        int p = left + up - upLeft;
        int pa = Math.Abs(p - left);
        int pb = Math.Abs(p - up);
        int pc = Math.Abs(p - upLeft);
        if (pa <= pb && pa <= pc) return left;
        return pb <= pc ? up : upLeft;
    }
}
