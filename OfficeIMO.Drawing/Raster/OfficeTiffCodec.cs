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

    private static void ValidateOptions(OfficeTiffEncodeOptions options) {
        if (options.Compression != OfficeTiffCompression.None && options.Compression != OfficeTiffCompression.PackBits) {
            throw new ArgumentOutOfRangeException(nameof(options.Compression));
        }

        ValidateDpi(options.DpiX, nameof(options.DpiX));
        ValidateDpi(options.DpiY, nameof(options.DpiY));
    }

    private static void ValidateDpi(double dpi, string name) {
        if (dpi <= 0D || double.IsNaN(dpi) || double.IsInfinity(dpi) || dpi > 1000000D) {
            throw new ArgumentOutOfRangeException(name, "DPI must be finite, positive, and no greater than 1,000,000.");
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
