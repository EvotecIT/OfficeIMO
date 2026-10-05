using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // Floating TIFF samples are normalized device components, not an implicit scRGB
    // declaration. Preserve their precision for unassociation and explicit ICC conversion.
    private static void ConvertFloatingPixel(byte[] source, int offset, int sampleBytes, bool littleEndian,
        int photometric, int alphaIndex, int alphaKind, double[]? colorComponents,
        out byte red, out byte green, out byte blue, out byte alpha) {
        double alphaSample = alphaIndex >= 0 ? ReadFloatingSample(source, offset + alphaIndex * sampleBytes, sampleBytes, littleEndian) : 1D;
        alpha = QuantizeFloatingComponent(alphaSample);
        double Component(int channel) {
            double value = ReadFloatingSample(source, offset + channel * sampleBytes, sampleBytes, littleEndian);
            if (alphaKind == 1) value = alphaSample <= 0D ? 0D : value / alphaSample;
            // Device ICC inputs are normalized; clipping also bounds division overflow
            // when finite HDR color is associated with a subnormal positive alpha.
            return Math.Max(0D, Math.Min(1D, value));
        }
        if (photometric == 0 || photometric == 1) {
            double gray = Component(0);
            if (photometric == 0) gray = 1D - gray;
            if (colorComponents != null) colorComponents[0] = gray;
            red = green = blue = QuantizeFloatingComponent(gray);
        } else if (photometric == 2) {
            double r = Component(0), g = Component(1), b = Component(2);
            if (colorComponents != null) { colorComponents[0] = r; colorComponents[1] = g; colorComponents[2] = b; }
            red = QuantizeFloatingComponent(r); green = QuantizeFloatingComponent(g); blue = QuantizeFloatingComponent(b);
        } else {
            double c = Component(0), m = Component(1), y = Component(2), k = Component(3);
            if (colorComponents != null) {
                colorComponents[0] = c; colorComponents[1] = m; colorComponents[2] = y; colorComponents[3] = k;
            }
            int black = QuantizeFloatingComponent(k);
            red = (byte)(255 - Math.Min(255, QuantizeFloatingComponent(c) + black));
            green = (byte)(255 - Math.Min(255, QuantizeFloatingComponent(m) + black));
            blue = (byte)(255 - Math.Min(255, QuantizeFloatingComponent(y) + black));
        }
    }

    private static byte QuantizeFloatingComponent(double value) =>
        (byte)Math.Floor(Math.Max(0D, Math.Min(1D, value)) * 255D + 0.5D);

    private static double ReadFloatingSample(byte[] source, int offset, int sampleBytes, bool littleEndian) {
        ulong bits = 0;
        for (int i = 0; i < sampleBytes; i++) bits = (bits << 8) | source[offset + (littleEndian ? sampleBytes - 1 - i : i)];
        double value;
        if (sampleBytes == 8) value = BitConverter.Int64BitsToDouble(unchecked((long)bits));
        else {
            // TN3 float24 has one sign, seven exponent and sixteen fraction bits.
            // 0x3F0000 is 1 and 0x000001 is 2^-78.
            int mantissaBits = sampleBytes == 2 ? 10 : sampleBytes == 3 ? 16 : 23;
            int bias = sampleBytes == 2 ? 15 : sampleBytes == 3 ? 63 : 127;
            int exponentMask = sampleBytes == 2 ? 31 : sampleBytes == 3 ? 127 : 255;
            int exponent = (int)(bits >> mantissaBits) & exponentMask;
            if (exponent == exponentMask) throw new FormatException("Non-finite TIFF samples cannot form SDR pixels.");
            ulong mantissa = bits & ((1UL << mantissaBits) - 1);
            value = exponent == 0 ? mantissa * Math.Pow(2D, 1 - bias - mantissaBits)
                : ((1UL << mantissaBits) + mantissa) * Math.Pow(2D, exponent - bias - mantissaBits);
            if ((bits & (1UL << (sampleBytes * 8 - 1))) != 0) value = -value;
        }
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new FormatException("Non-finite TIFF samples cannot form SDR pixels.");
        return value;
    }

    private static void ValidateFloatingSamples(byte[] bytes, int offset, int storedWidth, int columns, int rows,
        int samples, int baseSamples, int alphaIndex, int sampleBytes, bool littleEndian, CancellationToken cancellation) {
        for (int y = 0; y < rows; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < columns; x++) {
                if ((x & 0xFFF) == 0) cancellation.ThrowIfCancellationRequested();
                int pixel = offset + (y * storedWidth + x) * samples * sampleBytes;
                for (int c = 0; c < baseSamples; c++)
                    ReadFloatingSample(bytes, pixel + c * sampleBytes, sampleBytes, littleEndian);
                if (alphaIndex >= 0) ReadFloatingSample(bytes, pixel + alphaIndex * sampleBytes, sampleBytes, littleEndian);
            }
        }
    }

    // Technical Note 3: accumulate shuffled bytes at the component stride, then
    // restore sample bytes from most-significant-byte planes to the file byte order.
    private static void ReverseFloatingPredictor(byte[] bytes, int offset, int rows, int width,
        int samples, int sampleBytes, bool littleEndian, CancellationToken cancellation) {
        int sampleCount = checked(width * samples), rowBytes = checked(sampleCount * sampleBytes);
        var row = new byte[rowBytes];
        for (int y = 0; y < rows; y++) {
            cancellation.ThrowIfCancellationRequested();
            int start = checked(offset + y * rowBytes);
            Buffer.BlockCopy(bytes, start, row, 0, rowBytes);
            for (int i = samples; i < rowBytes; i++) {
                if ((i & 0x3FFF) == 0) cancellation.ThrowIfCancellationRequested();
                row[i] = unchecked((byte)(row[i] + row[i - samples]));
            }
            for (int i = 0; i < sampleCount; i++) {
                if ((i & 0xFFF) == 0) cancellation.ThrowIfCancellationRequested();
                for (int b = 0; b < sampleBytes; b++)
                    bytes[start + i * sampleBytes + (littleEndian ? sampleBytes - 1 - b : b)] = row[b * sampleCount + i];
            }
        }
    }
}
