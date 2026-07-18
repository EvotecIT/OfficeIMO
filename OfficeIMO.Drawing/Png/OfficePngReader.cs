using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free PNG decoder for common non-interlaced PNG images.
/// </summary>
public static class OfficePngReader {
    private static readonly byte[] Signature = { 137, 80, 78, 71, 13, 10, 26, 10 };

    /// <summary>
    /// Attempts to decode a PNG image into an RGBA raster buffer.
    /// </summary>
    public static bool TryDecode(byte[] bytes, out OfficeRasterImage? image) {
        image = null;
        try {
            if (bytes == null || !HasSignature(bytes)) {
                return false;
            }
            OfficeRasterGuards.EnsurePayloadWithinLimits(bytes.Length, "PNG payload exceeds size limits.");

            int width = 0;
            int height = 0;
            int bitDepth = 0;
            int colorType = 0;
            int compressionMethod = 0;
            int filterMethod = 0;
            int interlaceMethod = 0;
            byte[]? palette = null;
            byte[]? transparency = null;
            using MemoryStream idat = new MemoryStream();
            int offset = Signature.Length;
            while (offset + 12 <= bytes.Length) {
                int length = ReadBigEndianInt32(bytes, offset);
                long chunkEnd = (long)offset + 12L + length;
                if (length < 0 || chunkEnd > bytes.Length) {
                    return false;
                }

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
                    idat.Write(bytes, dataOffset, length);
                } else if (type == "IEND") {
                    break;
                }

                offset = (int)chunkEnd;
            }

            if (width <= 0 || height <= 0 || compressionMethod != 0 || filterMethod != 0 || interlaceMethod != 0 ||
                !IsSupportedColorLayout(colorType, bitDepth, palette)) {
                return false;
            }
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, out _)) return false;

            int bitsPerPixel = GetBitsPerPixel(colorType, bitDepth);
            int bytesPerPixel = Math.Max(1, (bitsPerPixel + 7) / 8);
            byte[] compressed = idat.ToArray();
            if (compressed.Length < 6) {
                return false;
            }

            using MemoryStream source = new MemoryStream(compressed, 2, compressed.Length - 6);
            using DeflateStream deflate = new DeflateStream(source, CompressionMode.Decompress);
            int stride = OfficeRasterGuards.EnsureByteCount(
                (((long)width * bitsPerPixel) + 7L) / 8L,
                "PNG scanline dimensions exceed size limits.");
            int expectedScanlineBytes = OfficeRasterGuards.EnsureByteCount(
                (long)(stride + 1) * height,
                "PNG decompressed data exceeds size limits.");
            byte[] scanlines = new byte[expectedScanlineBytes];
            int inflatedOffset = 0;
            while (inflatedOffset < scanlines.Length) {
                int read = deflate.Read(scanlines, inflatedOffset, scanlines.Length - inflatedOffset);
                if (read <= 0) return false;
                inflatedOffset += read;
            }
            // A valid non-interlaced PNG has exactly one filter byte plus one row payload per row.
            // Reject trailing decompressed data rather than allowing compressed input to inflate without bound.
            if (deflate.ReadByte() != -1) return false;
            byte[] previous = new byte[stride];
            byte[] current = new byte[stride];
            OfficeRasterImage result = new OfficeRasterImage(width, height);
            int sourceOffset = 0;
            for (int y = 0; y < height; y++) {
                if (sourceOffset >= scanlines.Length) return false;
                int filter = scanlines[sourceOffset++];
                if (sourceOffset + stride > scanlines.Length) return false;
                Buffer.BlockCopy(scanlines, sourceOffset, current, 0, stride);
                sourceOffset += stride;
                Unfilter(current, previous, bytesPerPixel, filter);
                ExpandScanline(current, width, y, colorType, bitDepth, palette, transparency, result);

                byte[] temp = previous;
                previous = current;
                current = temp;
                Array.Clear(current, 0, current.Length);
            }

            image = result;
            return true;
        } catch {
            image = null;
            return false;
        }
    }

    private static bool HasSignature(byte[] bytes) {
        if (bytes.Length < Signature.Length) return false;
        for (int i = 0; i < Signature.Length; i++) {
            if (bytes[i] != Signature[i]) return false;
        }

        return true;
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

    private static void ExpandScanline(byte[] current, int width, int y, int colorType, int bitDepth, byte[]? palette, byte[]? transparency, OfficeRasterImage image) {
        for (int x = 0; x < width; x++) {
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

            image.SetPixel(x, y, color);
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

    private static void Unfilter(byte[] current, byte[] previous, int bytesPerPixel, int filter) {
        for (int i = 0; i < current.Length; i++) {
            int left = i >= bytesPerPixel ? current[i - bytesPerPixel] : 0;
            int up = previous[i];
            int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
            int value = current[i];
            switch (filter) {
                case 0:
                    break;
                case 1:
                    value += left;
                    break;
                case 2:
                    value += up;
                    break;
                case 3:
                    value += (left + up) / 2;
                    break;
                case 4:
                    value += Paeth(left, up, upLeft);
                    break;
                default:
                    throw new InvalidDataException("Unsupported PNG filter.");
            }

            current[i] = (byte)(value & 0xFF);
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
