using System;
using System.IO;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.Drawing;

public static partial class OfficePngReader {
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

        if (colorType == 2 && bitDepth == 8 && transparency == null
            && destinationStartX == 0 && destinationStepX == 1) {
            ExpandRgb8Row(current, image.PixelBuffer, checked(y * image.Width * 4), width, cancellationToken);
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

    private static void ExpandRgb8Row(byte[] rgb, byte[] rgba, int destinationOffset, int width,
        CancellationToken cancellationToken) {
        for (int blockStart = 0; blockStart < width;) {
            cancellationToken.ThrowIfCancellationRequested();
            int blockEnd = blockStart + Math.Min(4096, width - blockStart);
            int pixel = blockStart;
#if NET8_0_OR_GREATER
            if (Ssse3.IsSupported) {
                var shuffle = Vector128.Create((byte)0, 1, 2, 128, 3, 4, 5, 128, 6, 7, 8, 128, 9, 10, 11, 128);
                var alpha = Vector128.Create(0xFF000000U).AsByte();
                // Four pixels consume twelve bytes. A sixteen-byte load must
                // also stay inside the row, including its unused high bytes.
                for (; pixel <= blockEnd - 4 && pixel * 3 <= rgb.Length - 16; pixel += 4) {
                    var source = Vector128.LoadUnsafe(ref rgb[pixel * 3]);
                    Sse2.Or(Ssse3.Shuffle(source, shuffle), alpha).StoreUnsafe(ref rgba[destinationOffset + pixel * 4]);
                }
            }
#endif
            for (; pixel < blockEnd; pixel++) {
                int source = pixel * 3;
                int destination = destinationOffset + pixel * 4;
                rgba[destination] = rgb[source];
                rgba[destination + 1] = rgb[source + 1];
                rgba[destination + 2] = rgb[source + 2];
                rgba[destination + 3] = 255;
            }
            blockStart = blockEnd;
        }
    }
}
