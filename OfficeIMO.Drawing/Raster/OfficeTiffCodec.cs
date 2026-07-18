using System;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>TIFF compression methods supported by the dependency-free encoder.</summary>
public enum OfficeTiffCompression {
    /// <summary>Writes uncompressed RGBA strips.</summary>
    None = 1,

    /// <summary>Uses the TIFF PackBits run-length encoding.</summary>
    PackBits = 32773
}

/// <summary>Settings for baseline TIFF encoding.</summary>
public sealed class OfficeTiffEncodeOptions {
    /// <summary>Strip compression. PackBits is dependency-free and broadly supported.</summary>
    public OfficeTiffCompression Compression { get; set; } = OfficeTiffCompression.PackBits;

    /// <summary>Horizontal resolution in dots per inch.</summary>
    public double DpiX { get; set; } = 96D;

    /// <summary>Vertical resolution in dots per inch.</summary>
    public double DpiY { get; set; } = 96D;
}

/// <summary>
/// Dependency-free baseline TIFF encoder for single-page RGBA images.
/// </summary>
public static class OfficeTiffCodec {
    private const int EntryCount = 15;

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

        byte[] pixels = image.GetPixels();
        byte[] strip = effective.Compression == OfficeTiffCompression.PackBits
            ? EncodePackBits(pixels)
            : pixels;

        const int ifdOffset = 8;
        int ifdLength = 2 + (EntryCount * 12) + 4;
        int bitsPerSampleOffset = checked(ifdOffset + ifdLength);
        int xResolutionOffset = checked(bitsPerSampleOffset + 8);
        int yResolutionOffset = checked(xResolutionOffset + 8);
        int stripOffset = checked(yResolutionOffset + 8);
        int fileLength = checked(stripOffset + strip.Length);
        byte[] output = new byte[fileLength];

        output[0] = (byte)'I';
        output[1] = (byte)'I';
        WriteUInt16(output, 2, 42);
        WriteUInt32(output, 4, ifdOffset);
        WriteUInt16(output, ifdOffset, EntryCount);

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
        WriteEntry(output, ref entry, 279, 4, 1, strip.Length);
        WriteEntry(output, ref entry, 282, 5, 1, xResolutionOffset);
        WriteEntry(output, ref entry, 283, 5, 1, yResolutionOffset);
        WriteShortEntry(output, ref entry, 284, 1);
        WriteShortEntry(output, ref entry, 296, 2);
        WriteShortEntry(output, ref entry, 338, 2);
        WriteUInt32(output, entry, 0);

        WriteUInt16(output, bitsPerSampleOffset, 8);
        WriteUInt16(output, bitsPerSampleOffset + 2, 8);
        WriteUInt16(output, bitsPerSampleOffset + 4, 8);
        WriteUInt16(output, bitsPerSampleOffset + 6, 8);
        WriteRational(output, xResolutionOffset, effective.DpiX);
        WriteRational(output, yResolutionOffset, effective.DpiY);
        Buffer.BlockCopy(strip, 0, output, stripOffset, strip.Length);
        return output;
    }

    /// <summary>
    /// Attempts to decode a classic baseline chunky RGB/RGBA TIFF using uncompressed or PackBits strips.
    /// BigTIFF, palette, tiled, planar, CMYK, floating-point, and compressed photographic variants are
    /// intentionally left to an optional caller codec.
    /// </summary>
    public static bool TryDecode(byte[]? encodedBytes, out OfficeRasterImage? image) {
        image = null;
        if (!IsTiff(encodedBytes) || encodedBytes == null ||
            encodedBytes.Length > OfficeRasterGuards.MaximumEncodedBytes) {
            return false;
        }

        try {
            bool littleEndian = encodedBytes[0] == (byte)'I';
            if (ReadUInt16(encodedBytes, 2, littleEndian) != 42) return false;
            int ifdOffset = ReadOffset(encodedBytes, 4, littleEndian);
            if (!HasBytes(encodedBytes, ifdOffset, 2)) return false;
            int entryCount = ReadUInt16(encodedBytes, ifdOffset, littleEndian);
            if (entryCount <= 0 || !HasBytes(encodedBytes, ifdOffset + 2, checked(entryCount * 12 + 4))) return false;

            var entries = new System.Collections.Generic.Dictionary<int, TiffEntry>();
            int entryOffset = ifdOffset + 2;
            for (int index = 0; index < entryCount; index++, entryOffset += 12) {
                int tag = ReadUInt16(encodedBytes, entryOffset, littleEndian);
                int type = ReadUInt16(encodedBytes, entryOffset + 2, littleEndian);
                uint count = ReadUInt32(encodedBytes, entryOffset + 4, littleEndian);
                if (count == 0 || count > int.MaxValue || entries.ContainsKey(tag)) return false;
                entries.Add(tag, new TiffEntry(type, (int)count, entryOffset + 8));
            }

            if (!TryReadScalar(encodedBytes, entries, 256, littleEndian, out int width) ||
                !TryReadScalar(encodedBytes, entries, 257, littleEndian, out int height) ||
                !OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) {
                return false;
            }

            if (!TryReadScalarOrDefault(encodedBytes, entries, 259, littleEndian, 1, out int compression) ||
                !TryReadScalarOrDefault(encodedBytes, entries, 262, littleEndian, 2, out int photometric) ||
                !TryReadScalarOrDefault(encodedBytes, entries, 274, littleEndian, 1, out int orientation) ||
                !TryReadScalarOrDefault(encodedBytes, entries, 277, littleEndian, 3, out int samples) ||
                !TryReadScalarOrDefault(encodedBytes, entries, 278, littleEndian, height, out int rowsPerStrip) ||
                !TryReadScalarOrDefault(encodedBytes, entries, 284, littleEndian, 1, out int planarConfiguration)) {
                return false;
            }
            if ((compression != (int)OfficeTiffCompression.None && compression != (int)OfficeTiffCompression.PackBits) ||
                photometric != 2 ||
                orientation < 1 || orientation > 4 ||
                (samples != 3 && samples != 4) ||
                rowsPerStrip < 1 ||
                planarConfiguration != 1) {
                return false;
            }

            int expectedStripCount = checked((height + rowsPerStrip - 1) / rowsPerStrip);
            if (!TryReadValues(encodedBytes, entries, 258, littleEndian, samples, out int[] bitsPerSample) ||
                Array.Exists(bitsPerSample, value => value != 8) ||
                !TryReadValues(encodedBytes, entries, 273, littleEndian, expectedStripCount, out int[] stripOffsets) ||
                !TryReadValues(encodedBytes, entries, 279, littleEndian, expectedStripCount, out int[] stripByteCounts)) {
                return false;
            }

            int alphaKind = 2;
            if (samples == 4) {
                if (!TryReadValues(encodedBytes, entries, 338, littleEndian, 1, out int[] extraSamples) ||
                    (extraSamples[0] != 1 && extraSamples[0] != 2)) {
                    return false;
                }
                alphaKind = extraSamples[0];
            }

            int sourceLength = OfficeRasterGuards.EnsureByteCount(
                (long)width * height * samples,
                "TIFF decoded source pixels exceed the managed limit.");
            byte[] source = new byte[sourceLength];
            int destinationOffset = 0;
            for (int strip = 0; strip < stripOffsets.Length && destinationOffset < source.Length; strip++) {
                int rowStart = checked(strip * rowsPerStrip);
                if (rowStart >= height) return false;
                int rows = Math.Min(rowsPerStrip, height - rowStart);
                int expected = checked(rows * width * samples);
                int offset = stripOffsets[strip];
                int count = stripByteCounts[strip];
                if (count < 0 || !HasBytes(encodedBytes, offset, count)) return false;
                bool decoded = compression == (int)OfficeTiffCompression.None
                    ? CopyExact(encodedBytes, offset, count, source, destinationOffset, expected)
                    : TryDecodePackBits(encodedBytes, offset, count, source, destinationOffset, expected);
                if (!decoded) return false;
                destinationOffset += expected;
            }
            if (destinationOffset != source.Length) return false;

            byte[] rgba = OfficeRasterGuards.AllocateRgba32(width, height, "TIFF decoded pixels exceed the managed limit.");
            for (int y = 0; y < height; y++) {
                for (int x = 0; x < width; x++) {
                    int sourcePixel = ((y * width) + x) * samples;
                    int targetX = orientation == 2 || orientation == 3 ? width - 1 - x : x;
                    int targetY = orientation == 3 || orientation == 4 ? height - 1 - y : y;
                    int targetPixel = ((targetY * width) + targetX) * 4;
                    byte alpha = samples == 4 ? source[sourcePixel + 3] : (byte)255;
                    rgba[targetPixel] = samples == 4 && alphaKind == 1
                        ? Unpremultiply(source[sourcePixel], alpha)
                        : source[sourcePixel];
                    rgba[targetPixel + 1] = samples == 4 && alphaKind == 1
                        ? Unpremultiply(source[sourcePixel + 1], alpha)
                        : source[sourcePixel + 1];
                    rgba[targetPixel + 2] = samples == 4 && alphaKind == 1
                        ? Unpremultiply(source[sourcePixel + 2], alpha)
                        : source[sourcePixel + 2];
                    rgba[targetPixel + 3] = alpha;
                }
            }

            image = OfficeRasterImage.FromRgba32(width, height, rgba);
            return true;
        } catch (ArgumentException) {
            return false;
        } catch (FormatException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
    }

    private static void ValidateOptions(OfficeTiffEncodeOptions options) {
        if (options.Compression != OfficeTiffCompression.None && options.Compression != OfficeTiffCompression.PackBits) {
            throw new ArgumentOutOfRangeException(nameof(options.Compression));
        }

        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
    }

    private static void ValidateDpi(double dpi, string name) {
        if (dpi < OfficeRasterImageEncoder.TiffMinimumDpi ||
            double.IsNaN(dpi) ||
            double.IsInfinity(dpi) ||
            dpi > 1000000D) {
            throw new ArgumentOutOfRangeException(name, "TIFF DPI must be finite and between 0.001 and 1,000,000.");
        }
    }

    private static byte[] EncodePackBits(byte[] input) {
        using var output = new MemoryStream(input.Length);
        int index = 0;
        while (index < input.Length) {
            int runLength = CountRun(input, index);
            if (runLength >= 3) {
                output.WriteByte(unchecked((byte)(257 - runLength)));
                output.WriteByte(input[index]);
                index += runLength;
                continue;
            }

            int literalStart = index;
            int literalLength = 0;
            while (index < input.Length && literalLength < 128) {
                runLength = CountRun(input, index);
                if (runLength >= 3) break;
                int take = Math.Min(runLength, 128 - literalLength);
                index += take;
                literalLength += take;
            }

            output.WriteByte((byte)(literalLength - 1));
            output.Write(input, literalStart, literalLength);
        }

        return output.ToArray();
    }

    private static int CountRun(byte[] input, int index) {
        int length = 1;
        while (length < 128 && index + length < input.Length && input[index + length] == input[index]) {
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
        int expectedCount) {
        int inputEnd = checked(inputOffset + inputCount);
        int outputEnd = checked(outputOffset + expectedCount);
        int source = inputOffset;
        int target = outputOffset;
        while (source < inputEnd && target < outputEnd) {
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
        while (source < inputEnd) {
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
        int expectedCount) {
        if (inputCount != expectedCount) return false;
        Buffer.BlockCopy(input, inputOffset, output, outputOffset, expectedCount);
        return true;
    }

    private static byte Unpremultiply(byte value, byte alpha) {
        if (alpha == 0) return 0;
        return (byte)Math.Min(255, (value * 255 + alpha / 2) / alpha);
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
