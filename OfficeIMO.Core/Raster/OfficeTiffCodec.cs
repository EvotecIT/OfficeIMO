using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using OfficeIMO.Core.Internal;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>TIFF compression methods supported by the dependency-free encoder.</summary>
public enum OfficeTiffCompression {
    /// <summary>Writes uncompressed RGBA strips.</summary>
    None = 1,

    /// <summary>Uses the TIFF 6.0 LZW dictionary compression method.</summary>
    Lzw = 5,

    /// <summary>Uses the TIFF PackBits run-length encoding.</summary>
    PackBits = 32773,

    /// <summary>Uses the Adobe Deflate compression method.</summary>
    Deflate = 8
}

/// <summary>Reversible sample predictor applied before TIFF compression.</summary>
public enum OfficeTiffPredictor {
    /// <summary>Compresses original sample bytes.</summary>
    None = 1,
    /// <summary>Stores horizontal sample differences before LZW or Deflate compression.</summary>
    Horizontal = 2
}

/// <summary>Settings for bounded classic TIFF encoding.</summary>
public sealed class OfficeTiffEncodeOptions {
    /// <summary>Strip compression. PackBits is dependency-free and broadly supported.</summary>
    public OfficeTiffCompression Compression { get; set; } = OfficeTiffCompression.PackBits;

    /// <summary>Predictor used with LZW or Deflate. Other compression modes always store original samples.</summary>
    public OfficeTiffPredictor Predictor { get; set; } = OfficeTiffPredictor.Horizontal;

    /// <summary>Horizontal resolution in dots per inch.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical resolution in dots per inch.</summary>
    public double DpiY { get; set; } = 96D;

    /// <summary>Writes TIFF XResolution, YResolution, and ResolutionUnit tags.</summary>
    public bool WriteResolution { get; set; } = true;
}

/// <summary>
/// Dependency-free classic TIFF encoder for single-page RGBA images.
/// </summary>
public static partial class OfficeTiffCodec {
    private const int BaseEntryCount = 15;
    private const int MaximumIfdCount = 65535;

    /// <summary>Returns whether the payload starts with a TIFF byte-order marker and magic value.</summary>
    public static bool IsTiff(byte[]? encodedBytes) =>
        encodedBytes != null && encodedBytes.Length >= 4 &&
        ((encodedBytes[0] == (byte)'I' && encodedBytes[1] == (byte)'I' && encodedBytes[2] == 42 && encodedBytes[3] == 0) ||
         (encodedBytes[0] == (byte)'M' && encodedBytes[1] == (byte)'M' && encodedBytes[2] == 0 && encodedBytes[3] == 42));

    /// <summary>Encodes a single RGBA image as a little-endian baseline TIFF.</summary>
    public static byte[] Encode(OfficeRasterImage image, OfficeTiffEncodeOptions? options = null) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        OfficeTiffEncodeOptions effective = options ?? new OfficeTiffEncodeOptions();
        ValidateOptions(effective);

        byte[] pixels = image.PixelBuffer;
        EnsureSinglePageCompressionWorkingSet(pixels.LongLength, effective);
        byte[]? strip = effective.Compression switch {
            OfficeTiffCompression.Deflate => OfficeZlibCodec.Compress(
                PrepareTiffCompressionInput(pixels, image.Width, image.Height, effective)),
            OfficeTiffCompression.Lzw => EncodeTiffLzw(pixels, image.Width, image.Height, effective),
            OfficeTiffCompression.None => pixels,
            _ => null
        };
        int stripLength = effective.Compression == OfficeTiffCompression.PackBits
            ? EncodePackBitsRows(pixels, image.Width * 4, image.Height, output: null, outputOffset: 0)
            : strip!.Length;

        const int ifdOffset = 8;
        bool writePredictor = UsesHorizontalPredictor(effective);
        int entryCount = BaseEntryCount - (effective.WriteResolution ? 0 : 3) + (writePredictor ? 1 : 0);
        int ifdLength = 2 + (entryCount * 12) + 4;
        int bitsPerSampleOffset = checked(ifdOffset + ifdLength);
        int xResolutionOffset = checked(bitsPerSampleOffset + 8);
        int yResolutionOffset = checked(xResolutionOffset + 8);
        int stripOffset = effective.WriteResolution ? checked(yResolutionOffset + 8) : xResolutionOffset;
        int fileLength = OfficeRasterGuards.EnsureOutputBytes(
            checked((long)stripOffset + stripLength),
            "The TIFF exceeds the encoded-size limit.");
        long retainedStripBytes = effective.Compression is OfficeTiffCompression.Lzw or OfficeTiffCompression.Deflate
            ? stripLength
            : 0L;
        if (!IsMultiPageTiffWorkingSetWithinLimit(pixels.LongLength, retainedStripBytes, fileLength)) {
            throw new ArgumentException("The TIFF encoding working set exceeds the managed limit.", nameof(image));
        }
        byte[] output = new byte[fileLength];

        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, entryCount);

        int entry = ifdOffset + 2;
        WriteEntry(output, ref entry, 256, 4, 1, image.Width);
        WriteEntry(output, ref entry, 257, 4, 1, image.Height);
        WriteEntry(output, ref entry, 258, 3, 4, bitsPerSampleOffset);
        WriteShortEntry(output, ref entry, 259, (int)effective.Compression);
        WriteShortEntry(output, ref entry, 262, 2);
        WriteEntry(output, ref entry, 273, 4, 1, stripOffset);
        WriteShortEntry(output, ref entry, 274, 1);
        WriteShortEntry(output, ref entry, 277, 4);
        WriteEntry(output, ref entry, 278, 4, 1, image.Height);
        WriteEntry(output, ref entry, 279, 4, 1, stripLength);
        if (effective.WriteResolution) {
            WriteEntry(output, ref entry, 282, 5, 1, xResolutionOffset);
            WriteEntry(output, ref entry, 283, 5, 1, yResolutionOffset);
        }
        WriteShortEntry(output, ref entry, 284, 1);
        if (effective.WriteResolution) WriteShortEntry(output, ref entry, 296, 2);
        if (writePredictor) WriteShortEntry(output, ref entry, 317, (int)effective.Predictor);
        WriteShortEntry(output, ref entry, 338, 2);
        WriteUInt32(output, entry, 0);

        WriteUInt16(output, bitsPerSampleOffset, 8);
        WriteUInt16(output, bitsPerSampleOffset + 2, 8);
        WriteUInt16(output, bitsPerSampleOffset + 4, 8);
        WriteUInt16(output, bitsPerSampleOffset + 6, 8);
        if (effective.WriteResolution) {
            WriteRational(output, xResolutionOffset, effective.DpiX);
            WriteRational(output, yResolutionOffset, effective.DpiY);
        }
        if (effective.Compression == OfficeTiffCompression.PackBits) {
            int written = EncodePackBitsRows(pixels, image.Width * 4, image.Height, output, stripOffset);
            if (written != stripLength) {
                throw new InvalidOperationException("The TIFF PackBits length calculation is inconsistent.");
            }
        } else {
            Buffer.BlockCopy(strip!, 0, output, stripOffset, stripLength);
        }
        return output;
    }

    /// <summary>Reads the number of top-level images from a bounded classic TIFF IFD chain.</summary>
    public static bool TryGetPageCount(byte[]? encodedBytes, out int pageCount) {
        pageCount = 0;
        if (!IsTiff(encodedBytes) || encodedBytes == null ||
            encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes) {
            return false;
        }

        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            if (ReadUInt16(encodedBytes, 2, littleEndian) != 42) return false;
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            int firstIfdOffset = ifdOffset;
            System.Collections.Generic.HashSet<int>? visitedIfds = null;
            while (ifdOffset != 0) {
                if (pageCount >= MaximumIfdCount || ifdOffset < 8 || !HasBytes(encodedBytes, ifdOffset, 2)) {
                    pageCount = 0;
                    return false;
                }
                if (pageCount > 0) {
                    visitedIfds ??= new System.Collections.Generic.HashSet<int> { firstIfdOffset };
                    if (!visitedIfds.Add(ifdOffset)) {
                        pageCount = 0;
                        return false;
                    }
                }
                int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
                if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) {
                    pageCount = 0;
                    return false;
                }
                pageCount++;
                int nextIfdPointerOffset = checked(ifdOffset + 2 + entryCount * 12);
                ifdOffset = ReadOffset(encodedBytes, nextIfdPointerOffset, littleEndian);
            }
            return pageCount > 0;
        } catch (ArgumentException) {
            pageCount = 0;
            return false;
        } catch (FormatException) {
            pageCount = 0;
            return false;
        } catch (OverflowException) {
            pageCount = 0;
            return false;
        }
    }

    /// <summary>
    /// Attempts to decode a classic baseline grayscale, palette, RGB, RGBA, or device-CMYK TIFF using
    /// chunky or planar strips or tiles with uncompressed, LZW, PackBits, or Deflate payloads.
    /// Floating-point, JPEG-compressed, and BigTIFF payloads remain optional caller-codec responsibilities.
    /// </summary>
    public static bool TryDecode(byte[]? encodedBytes, out OfficeRasterImage? image) =>
        TryDecodePage(encodedBytes, 0, options: null, out image);

    /// <summary>Attempts to decode one zero-based page from a bounded classic TIFF container.</summary>
    public static bool TryDecodePage(byte[]? encodedBytes, int pageIndex, out OfficeRasterImage? image) =>
        TryDecodePage(encodedBytes, pageIndex, options: null, out image);

    internal static bool TryDecodePage(
        byte[]? encodedBytes,
        int pageIndex,
        OfficeRasterDecodeOptions? options,
        out OfficeRasterImage? image) {
        image = null;
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        OfficeRasterDecodeOptions effective = options ?? new OfficeRasterDecodeOptions();
        effective.Validate();
        effective.CancellationToken.ThrowIfCancellationRequested();
        if (!IsTiff(encodedBytes) || encodedBytes == null ||
            encodedBytes.Length > effective.MaximumEncodedBytes ||
            !OfficeTiffStructureValidator.TryValidate(
                encodedBytes, 0, encodedBytes.Length, effective.CancellationToken)) {
            return false;
        }
        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            if (ReadUInt16(encodedBytes, 2, littleEndian) != 42) return false;
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            var visitedIfds = new System.Collections.Generic.HashSet<int>();
            int currentPageIndex = 0;
            while (ifdOffset != 0) {
                effective.CancellationToken.ThrowIfCancellationRequested();
                if (visitedIfds.Count >= MaximumIfdCount || !visitedIfds.Add(ifdOffset) ||
                    !HasBytes(encodedBytes, ifdOffset, 2)) {
                    return false;
                }
                int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
                if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) return false;

                var entries = new System.Collections.Generic.Dictionary<int, TiffEntry>();
                int entryOffset = ifdOffset + 2;
                for (int index = 0; index < entryCount; index++, entryOffset += 12) {
                    if ((index & 0xFF) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                    int tag = ReadUInt16(encodedBytes, entryOffset, littleEndian);
                    int type = ReadUInt16(encodedBytes, entryOffset + 2, littleEndian);
                    uint count = ReadUInt32(encodedBytes, entryOffset + 4, littleEndian);
                    if (count == 0 || count > int.MaxValue || entries.ContainsKey(tag) ||
                        !HasValidEntryValueRange(
                            encodedBytes,
                            type,
                            (int)count,
                            entryOffset + 8,
                            littleEndian)) {
                        return false;
                    }
                    entries.Add(tag, new TiffEntry(type, (int)count, entryOffset + 8));
                }

                int nextIfdPointerOffset = checked(ifdOffset + 2 + entryCount * 12);
                int nextIfdOffset = ReadOffset(encodedBytes, nextIfdPointerOffset, littleEndian);
                if (currentPageIndex != pageIndex) {
                    ifdOffset = nextIfdOffset;
                    currentPageIndex++;
                    continue;
                }

                if (!TryReadScalar(encodedBytes, entries, 256, littleEndian, out int width) ||
                    !TryReadScalar(encodedBytes, entries, 257, littleEndian, out int height) ||
                    !IsWithinPixelLimit(width, height, effective.MaximumDecodedPixels)) {
                    return false;
                }

                if (!TryReadScalarOrDefault(encodedBytes, entries, 259, littleEndian, 1, out int compression) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 262, littleEndian, 2, out int photometric) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out int orientation) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 278, littleEndian, height, out int rowsPerStrip) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 284, littleEndian, 1, out int planarConfiguration) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 317, littleEndian, 1, out int predictor)) {
                    return false;
                }
                if (!TryGetBaseSampleCount(photometric, out int baseSamples) ||
                    !TryReadScalarOrDefault(encodedBytes, entries, 277, littleEndian, baseSamples, out int samples)) {
                    return false;
                }
                if (photometric == 5 &&
                    (!TryReadScalarOrDefault(encodedBytes, entries, 332, littleEndian, 1, out int inkSet) ||
                     inkSet != 1)) {
                    return false;
                }
                if ((compression != (int)OfficeTiffCompression.None &&
                     compression != (int)OfficeTiffCompression.Lzw &&
                     compression != (int)OfficeTiffCompression.PackBits &&
                     compression != (int)OfficeTiffCompression.Deflate &&
                     compression != 32946) ||
                    orientation < 1 || orientation > 8 ||
                    (samples != baseSamples && samples != baseSamples + 1) ||
                    rowsPerStrip < 1 ||
                    (planarConfiguration != 1 && planarConfiguration != 2) ||
                    (predictor != 1 && predictor != 2)) {
                    return false;
                }

                if (!TryReadValues(encodedBytes, entries, 258, littleEndian, samples, out int[] bitsPerSample) ||
                    Array.Exists(bitsPerSample, value => value != 8)) {
                    return false;
                }

                int[]? colorMap = null;
                if (photometric == 3 &&
                    !TryReadValues(encodedBytes, entries, 320, littleEndian, 768, out colorMap)) {
                    return false;
                }

                int alphaKind = 2;
                if (samples == baseSamples + 1) {
                    if (!TryReadValues(encodedBytes, entries, 338, littleEndian, 1, out int[] extraSamples) ||
                        (extraSamples[0] != 1 && extraSamples[0] != 2)) {
                        return false;
                    }
                    alphaKind = extraSamples[0];
                }

                long maximumDecodeWorkBytes = OfficeRasterGuards.MaximumDecodedBytes - effective.RetainedManagedBytes;
                if (maximumDecodeWorkBytes < 1L) return false;
                var decodeWorkBudget = new TiffValidationBudget(maximumDecodeWorkBytes);
                if (!TryDecodePixelSegments(encodedBytes, entries, littleEndian, width, height, samples,
                        compression, planarConfiguration, predictor, effective, decodeWorkBudget,
                        retainPixels: true, out byte[] source)) return false;

                int orientedWidth = orientation >= 5 ? height : width;
                int orientedHeight = orientation >= 5 ? width : height;
                byte[] rgba = OfficeRasterGuards.AllocateRgba32(orientedWidth, orientedHeight, "TIFF decoded pixels exceed the managed limit.");
                for (int y = 0; y < height; y++) {
                    if ((y & 31) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                    for (int x = 0; x < width; x++) {
                        if ((x & 0xFFF) == 0) effective.CancellationToken.ThrowIfCancellationRequested();
                        int sourcePixel = ((y * width) + x) * samples;
                        ResolveOrientedPixel(x, y, width, height, orientation, out int targetX, out int targetY);
                        int targetPixel = ((targetY * orientedWidth) + targetX) * 4;
                        byte alpha = samples == baseSamples + 1
                            ? source[sourcePixel + baseSamples]
                            : (byte)255;
                        ConvertPixel(
                            source,
                            sourcePixel,
                            photometric,
                            alphaKind,
                            alpha,
                            colorMap,
                            out byte red,
                            out byte green,
                            out byte blue);
                        rgba[targetPixel] = red;
                        rgba[targetPixel + 1] = green;
                        rgba[targetPixel + 2] = blue;
                        rgba[targetPixel + 3] = alpha;
                    }
                }
                image = OfficeRasterImage.FromOwnedRgba32(orientedWidth, orientedHeight, rgba);
                return true;
            }
            return false;
        } catch (ArgumentException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static bool IsWithinPixelLimit(int width, int height, long maximumPixels) =>
        width > 0 && height > 0 && width <= maximumPixels && height <= maximumPixels / width;

    private static void ResolveOrientedPixel(
        int x,
        int y,
        int width,
        int height,
        int orientation,
        out int targetX,
        out int targetY) {
        switch (orientation) {
            case 2:
                targetX = width - 1 - x;
                targetY = y;
                break;
            case 3:
                targetX = width - 1 - x;
                targetY = height - 1 - y;
                break;
            case 4:
                targetX = x;
                targetY = height - 1 - y;
                break;
            case 5:
                targetX = y;
                targetY = x;
                break;
            case 6:
                targetX = height - 1 - y;
                targetY = x;
                break;
            case 7:
                targetX = height - 1 - y;
                targetY = width - 1 - x;
                break;
            case 8:
                targetX = y;
                targetY = width - 1 - x;
                break;
            default:
                targetX = x;
                targetY = y;
                break;
        }
    }

    private static void ValidateOptions(OfficeTiffEncodeOptions options) {
        if (options.Compression != OfficeTiffCompression.None &&
            options.Compression != OfficeTiffCompression.Lzw &&
            options.Compression != OfficeTiffCompression.PackBits &&
            options.Compression != OfficeTiffCompression.Deflate) {
            throw new ArgumentOutOfRangeException(nameof(options.Compression));
        }
        if (options.Predictor != OfficeTiffPredictor.None && options.Predictor != OfficeTiffPredictor.Horizontal) {
            throw new ArgumentOutOfRangeException(nameof(options.Predictor));
        }

        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
    }

    private static bool UsesHorizontalPredictor(OfficeTiffEncodeOptions options) =>
        options.Predictor == OfficeTiffPredictor.Horizontal &&
        (options.Compression == OfficeTiffCompression.Lzw || options.Compression == OfficeTiffCompression.Deflate);

    private static byte[] PrepareTiffCompressionInput(
        byte[] pixels,
        int width,
        int height,
        OfficeTiffEncodeOptions options) {
        if (!UsesHorizontalPredictor(options)) return pixels;
        byte[] predicted = (byte[])pixels.Clone();
        ApplyHorizontalPredictor(predicted, width, height);
        return predicted;
    }

    private static byte[] EncodeTiffLzw(
        byte[] pixels,
        int width,
        int height,
        OfficeTiffEncodeOptions options) {
        if (!UsesHorizontalPredictor(options)) return EncodeLzw(pixels, pixels.Length);
#if NET8_0_OR_GREATER
        byte[] scratch = ArrayPool<byte>.Shared.Rent(pixels.Length);
#else
        byte[] scratch = new byte[pixels.Length];
#endif
        try {
            Buffer.BlockCopy(pixels, 0, scratch, 0, pixels.Length);
            ApplyHorizontalPredictor(scratch, width, height);
            return EncodeLzw(scratch, pixels.Length);
        } finally {
#if NET8_0_OR_GREATER
            ArrayPool<byte>.Shared.Return(scratch);
#endif
        }
    }

    private static void ApplyHorizontalPredictor(byte[] pixels, int width, int height) {
        const int samples = 4;
        int rowBytes = checked(width * samples);
        for (int y = 0; y < height; y++) {
            int row = y * rowBytes;
            for (int offset = row + rowBytes - 1; offset >= row + samples; offset--) {
                pixels[offset] = unchecked((byte)(pixels[offset] - pixels[offset - samples]));
            }
        }
    }

    private static void ValidateDpi(double dpi, string name) {
        if (dpi < OfficeRasterImageEncoder.TiffMinimumDpi ||
            double.IsNaN(dpi) ||
            double.IsInfinity(dpi) ||
            dpi > OfficeRasterImageEncoder.TiffMaximumDpi) {
            throw new ArgumentOutOfRangeException(name, "TIFF DPI must be finite and between 0.001 and 1,000,000.");
        }
    }

    private static int EncodePackBitsRows(
        byte[] input,
        int rowBytes,
        int rowCount,
        byte[]? output,
        int outputOffset) {
        if (rowBytes <= 0 || rowCount <= 0 || (long)rowBytes * rowCount != input.Length) {
            throw new ArgumentException("TIFF PackBits row dimensions do not match the input buffer.");
        }
        int target = outputOffset;
        for (int row = 0; row < rowCount; row++) {
            target += EncodePackBits(input, row * rowBytes, rowBytes, output, target);
        }
        return checked(target - outputOffset);
    }

    private static int EncodePackBits(
        byte[] input,
        int inputOffset,
        int inputCount,
        byte[]? output,
        int outputOffset) {
        int index = inputOffset;
        int inputEnd = checked(inputOffset + inputCount);
        int target = outputOffset;
        while (index < inputEnd) {
            int runLength = CountRun(input, index, inputEnd);
            if (runLength >= 3) {
                if (output != null) {
                    output[target] = unchecked((byte)(257 - runLength));
                    output[target + 1] = input[index];
                }
                target += 2;
                index += runLength;
                continue;
            }

            int literalStart = index;
            int literalLength = 0;
            while (index < inputEnd && literalLength < 128) {
                runLength = CountRun(input, index, inputEnd);
                if (runLength >= 3) break;
                int take = Math.Min(runLength, 128 - literalLength);
                index += take;
                literalLength += take;
            }

            if (output != null) {
                output[target] = (byte)(literalLength - 1);
                Buffer.BlockCopy(input, literalStart, output, target + 1, literalLength);
            }
            target += literalLength + 1;
        }

        return checked(target - outputOffset);
    }

    private static int CountRun(byte[] input, int index, int inputEnd) {
        int length = 1;
        while (length < 128 && index + length < inputEnd && input[index + length] == input[index]) {
            length++;
        }

        return length;
    }

    private static bool TryDecodePackBits(
        byte[] input,
        int inputOffset,
        int inputCount,
        byte[] output,
        int outputOffset,
        int expectedCount,
        CancellationToken cancellationToken) {
        int inputEnd = checked(inputOffset + inputCount);
        int outputEnd = checked(outputOffset + expectedCount);
        int source = inputOffset;
        int target = outputOffset;
        while (source < inputEnd && target < outputEnd) {
            if (((source - inputOffset) & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int header = unchecked((sbyte)input[source++]);
            if (header >= 0) {
                int literalCount = header + 1;
                if (source > inputEnd - literalCount || target > outputEnd - literalCount) return false;
                Buffer.BlockCopy(input, source, output, target, literalCount);
                source += literalCount;
                target += literalCount;
            } else if (header >= -127) {
                int repeatCount = 1 - header;
                if (source >= inputEnd || target > outputEnd - repeatCount) return false;
                byte value = input[source++];
                for (int index = 0; index < repeatCount; index++) output[target++] = value;
            }
        }
        if (target != outputEnd) return false;
        return TryValidatePackBitsPadding(input, source, inputEnd - source, cancellationToken);
    }

    internal static bool TryValidatePackBitsPadding(
        byte[] input,
        int inputOffset,
        int inputCount,
        CancellationToken cancellationToken) {
        int inputEnd = checked(inputOffset + inputCount);
        int source = inputOffset;
        int trailingBytes = 0;
        while (source < inputEnd) {
            if ((trailingBytes++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (unchecked((sbyte)input[source++]) != -128) return false;
        }
        return true;
    }

    private static bool CopyExact(
        byte[] input,
        int inputOffset,
        int inputCount,
        byte[] output,
        int outputOffset,
        int expectedCount,
        CancellationToken cancellationToken) {
        if (inputCount != expectedCount) return false;
        CopyWithCancellation(input, inputOffset, output, outputOffset, expectedCount, cancellationToken);
        return true;
    }

    private static void CopyWithCancellation(
        byte[] input,
        int inputOffset,
        byte[] output,
        int outputOffset,
        int count,
        CancellationToken cancellationToken) {
        const int blockSize = 1024 * 1024;
        int copied = 0;
        while (copied < count) {
            cancellationToken.ThrowIfCancellationRequested();
            int current = Math.Min(blockSize, count - copied);
            Buffer.BlockCopy(input, inputOffset + copied, output, outputOffset + copied, current);
            copied += current;
        }
    }

    private static byte Unpremultiply(byte value, byte alpha) {
        if (alpha == 0) return 0;
        return (byte)Math.Min(255, (value * 255 + alpha / 2) / alpha);
    }

    private static bool TryGetBaseSampleCount(int photometric, out int samples) {
        samples = photometric switch {
            0 => 1,
            1 => 1,
            2 => 3,
            3 => 1,
            5 => 4,
            _ => 0
        };
        return samples != 0;
    }

    private static void ConvertPixel(
        byte[] source,
        int offset,
        int photometric,
        int alphaKind,
        byte alpha,
        int[]? colorMap,
        out byte red,
        out byte green,
        out byte blue) {
        byte Component(int componentOffset) =>
            alphaKind == 1
                ? Unpremultiply(source[offset + componentOffset], alpha)
                : source[offset + componentOffset];

        switch (photometric) {
            case 0:
            case 1:
                byte luminance = Component(0);
                if (photometric == 0) luminance = (byte)(255 - luminance);
                red = luminance;
                green = luminance;
                blue = luminance;
                return;
            case 2:
                red = Component(0);
                green = Component(1);
                blue = Component(2);
                return;
            case 3:
                int paletteIndex = source[offset];
                red = ColorMapByte(colorMap![paletteIndex]);
                green = ColorMapByte(colorMap[256 + paletteIndex]);
                blue = ColorMapByte(colorMap[512 + paletteIndex]);
                return;
            case 5:
                int cyan = Component(0);
                int magenta = Component(1);
                int yellow = Component(2);
                int black = Component(3);
                red = (byte)(255 - Math.Min(255, cyan + black));
                green = (byte)(255 - Math.Min(255, magenta + black));
                blue = (byte)(255 - Math.Min(255, yellow + black));
                return;
            default:
                red = green = blue = 0;
                return;
        }
    }

    private static byte ColorMapByte(int value) =>
        (byte)Math.Min(255, (value + 128) / 257);

    private static void ReverseHorizontalPredictor(
        byte[] pixels,
        int offset,
        int rows,
        int width,
        int samples,
        CancellationToken cancellationToken) {
        int rowBytes = checked(width * samples);
        for (int row = 0; row < rows; row++) {
            if ((row & 31) == 0) cancellationToken.ThrowIfCancellationRequested();
            int rowOffset = checked(offset + row * rowBytes);
            int rowEnd = checked(rowOffset + rowBytes);
            for (int index = rowOffset + samples; index < rowEnd; index++) {
                if (((index - rowOffset) & 0xFFF) == 0) cancellationToken.ThrowIfCancellationRequested();
                pixels[index] = unchecked((byte)(pixels[index] + pixels[index - samples]));
            }
        }
    }

    private static bool TryReadScalar(
        byte[] data,
        System.Collections.Generic.IReadOnlyDictionary<int, TiffEntry> entries,
        int tag,
        bool littleEndian,
        out int value) {
        value = 0;
        return TryReadValues(data, entries, tag, littleEndian, 1, out int[] values) &&
               (value = values[0]) >= 0;
    }

    private static bool TryReadScalarOrDefault(
        byte[] data,
        System.Collections.Generic.IReadOnlyDictionary<int, TiffEntry> entries,
        int tag,
        bool littleEndian,
        int defaultValue,
        out int value) {
        if (!entries.ContainsKey(tag)) {
            value = defaultValue;
            return true;
        }
        return TryReadScalar(data, entries, tag, littleEndian, out value);
    }

    private static bool TryReadValues(
        byte[] data,
        System.Collections.Generic.IReadOnlyDictionary<int, TiffEntry> entries,
        int tag,
        bool littleEndian,
        int expectedCount,
        out int[] values) =>
        TryReadValues(data, entries, tag, littleEndian, expectedCount,
            CancellationToken.None, out values);

    private static bool TryReadValues(
        byte[] data,
        System.Collections.Generic.IReadOnlyDictionary<int, TiffEntry> entries,
        int tag,
        bool littleEndian,
        int expectedCount,
        CancellationToken cancellationToken,
        out int[] values) {
        values = Array.Empty<int>();
        if (!entries.TryGetValue(tag, out TiffEntry entry) ||
            (entry.Type != 3 && entry.Type != 4) ||
            entry.Count != expectedCount) {
            return false;
        }
        int itemSize = entry.Type == 3 ? 2 : 4;
        int byteCount = checked(entry.Count * itemSize);
        int valueOffset = byteCount <= 4
            ? entry.ValueFieldOffset
            : ReadOffset(data, entry.ValueFieldOffset, littleEndian);
        if (!HasBytes(data, valueOffset, byteCount)) return false;
        values = new int[entry.Count];
        for (int index = 0; index < values.Length; index++) {
            if ((index & 0x3FF) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (entry.Type == 3) {
                values[index] = ReadUInt16(data, valueOffset + index * 2, littleEndian);
            } else {
                uint value = ReadUInt32(data, valueOffset + index * 4, littleEndian);
                if (value > int.MaxValue) return false;
                values[index] = (int)value;
            }
        }
        return true;
    }

    private static int ReadOffset(byte[] data, int offset, bool littleEndian) {
        uint value = ReadUInt32(data, offset, littleEndian);
        if (value > int.MaxValue) throw new FormatException("TIFF offset exceeds supported integer bounds.");
        return (int)value;
    }

    private static bool HasValidEntryValueRange(
        byte[] data,
        int type,
        int count,
        int valueFieldOffset,
        bool littleEndian) {
        int itemSize;
        switch (type) {
            case 1:  // BYTE
            case 2:  // ASCII
            case 6:  // SBYTE
            case 7:  // UNDEFINED
                itemSize = 1;
                break;
            case 3:  // SHORT
            case 8:  // SSHORT
                itemSize = 2;
                break;
            case 4:  // LONG
            case 9:  // SLONG
            case 11: // FLOAT
            case 13: // IFD
                itemSize = 4;
                break;
            case 5:  // RATIONAL
            case 10: // SRATIONAL
            case 12: // DOUBLE
                itemSize = 8;
                break;
            default:
                return false;
        }

        long byteCount = (long)count * itemSize;
        if (byteCount <= 4) return HasBytes(data, valueFieldOffset, 4);
        if (byteCount > int.MaxValue) return false;
        int valueOffset = ReadOffset(data, valueFieldOffset, littleEndian);
        return HasBytes(data, valueOffset, (int)byteCount);
    }

    private static int ReadUInt16(byte[] data, int offset, bool littleEndian) {
        if (!HasBytes(data, offset, 2)) throw new FormatException("TIFF field is truncated.");
        return littleEndian
            ? data[offset] | data[offset + 1] << 8
            : data[offset] << 8 | data[offset + 1];
    }

    private static uint ReadUInt32(byte[] data, int offset, bool littleEndian) {
        if (!HasBytes(data, offset, 4)) throw new FormatException("TIFF field is truncated.");
        return littleEndian
            ? (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24)
            : (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);
    }

    private static bool HasBytes(byte[] data, int offset, int count) =>
        offset >= 0 && count >= 0 && offset <= data.Length - count;

    private readonly struct TiffEntry {
        internal TiffEntry(int type, int count, int valueFieldOffset) {
            Type = type;
            Count = count;
            ValueFieldOffset = valueFieldOffset;
        }

        internal int Type { get; }
        internal int Count { get; }
        internal int ValueFieldOffset { get; }
    }

    private static void WriteEntry(byte[] output, ref int offset, int tag, int type, int count, int value) {
        WriteUInt16(output, offset, tag);
        WriteUInt16(output, offset + 2, type);
        WriteUInt32(output, offset + 4, count);
        WriteUInt32(output, offset + 8, value);
        offset += 12;
    }

    private static void WriteShortEntry(byte[] output, ref int offset, int tag, int value) {
        WriteUInt16(output, offset, tag);
        WriteUInt16(output, offset + 2, 3);
        WriteUInt32(output, offset + 4, 1);
        WriteUInt16(output, offset + 8, value);
        offset += 12;
    }

    private static void WriteRational(byte[] output, int offset, double value) {
        const int denominator = 1000;
        int numerator = checked((int)Math.Round(value * denominator));
        WriteUInt32(output, offset, numerator);
        WriteUInt32(output, offset + 4, denominator);
    }

    private static void WriteUInt16(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] output, int offset, int value) {
        output[offset] = (byte)value;
        output[offset + 1] = (byte)(value >> 8);
        output[offset + 2] = (byte)(value >> 16);
        output[offset + 3] = (byte)(value >> 24);
    }
}
