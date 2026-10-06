using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static int ReadTiffJpegSample(byte[] pixels, int sample, int sampleBytes, bool littleEndian) =>
        sampleBytes == 1 ? pixels[sample] : ReadUInt16(pixels, sample * 2, littleEndian);

    private static void WriteTiffJpegSample(byte[] pixels, int sample, int sampleBytes, bool littleEndian, int value) {
        int at = sample * sampleBytes;
        if (sampleBytes == 1) pixels[at] = (byte)value;
        else {
            pixels[at] = (byte)(littleEndian ? value : value >> 8);
            pixels[at + 1] = (byte)(littleEndian ? value >> 8 : value);
        }
    }

    // Preserve fractional RGB through unassociation and ICC conversion. Rounding
    // back to the JPEG precision loses visible color, especially at 2–7 bits.
    private sealed class TiffJpegColorTransform {
        private readonly int _maximum;
        private readonly double[] _coefficients;
        private readonly double[] _reference;

        internal byte[]? ChromaFractions { get; }

        internal TiffJpegColorTransform(int maximum, double[] coefficients, double[] reference, byte[]? chromaFractions) {
            _maximum = maximum;
            _coefficients = coefficients;
            _reference = reference;
            ChromaFractions = chromaFractions;
        }

        internal void ConvertPixel(byte[] pixels, int offset, int sampleBytes, int samples, bool littleEndian,
            int alphaIndex, int alphaKind, double[]? colorComponents,
            out byte red, out byte green, out byte blue, out byte alpha) {
            int sample = offset / sampleBytes;
            int fraction = sample / samples * 2;
            double cbFraction = ChromaFractions == null ? 0D : ChromaFractions[fraction] / 64D;
            double crFraction = ChromaFractions == null ? 0D : ChromaFractions[fraction + 1] / 64D;
            int a = alphaIndex < 0 ? _maximum : ReadTiffJpegSample(pixels, sample + alphaIndex, sampleBytes, littleEndian);
            alpha = QuantizeUnsigned16Component(a / (double)_maximum);
            double y = (ReadTiffJpegSample(pixels, sample, sampleBytes, littleEndian) - _reference[0]) * _maximum / (_reference[1] - _reference[0]);
            double cb = (ReadTiffJpegSample(pixels, sample + 1, sampleBytes, littleEndian) + cbFraction - _reference[2]) * (_maximum / 2) / (_reference[3] - _reference[2]);
            double cr = (ReadTiffJpegSample(pixels, sample + 2, sampleBytes, littleEndian) + crFraction - _reference[4]) * (_maximum / 2) / (_reference[5] - _reference[4]);
            double r = y + cr * (2 - 2 * _coefficients[0]);
            double b = y + cb * (2 - 2 * _coefficients[2]);
            double g = (y - _coefficients[0] * r - _coefficients[2] * b) / _coefficients[1];
            double Normalize(double value) => alphaKind == 1 && a == 0 ? 0D :
                Math.Max(0D, Math.Min(1D, value / (alphaKind == 1 ? a : _maximum)));
            r = Normalize(r); g = Normalize(g); b = Normalize(b);
            if (colorComponents != null) {
                colorComponents[0] = r; colorComponents[1] = g; colorComponents[2] = b;
            }
            red = QuantizeUnsigned16Component(r);
            green = QuantizeUnsigned16Component(g);
            blue = QuantizeUnsigned16Component(b);
        }
    }
}
