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

    private static void ConvertTiffJpegYcc(byte[] pixels, int samples, int sampleBytes,
        bool littleEndian, double[] c, double[] r, OfficeRasterDecodeOptions options) {
        int maximum = sampleBytes == 1 ? 255 : 65535;
        int Clamp(double value) => (int)Math.Round(Math.Max(0, Math.Min(maximum, value)));
        for (int i = 0; i < pixels.Length / sampleBytes; i += samples) {
            if ((i & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
            double y = (ReadTiffJpegSample(pixels, i, sampleBytes, littleEndian) - r[0]) * maximum / (r[1] - r[0]);
            double cb = (ReadTiffJpegSample(pixels, i + 1, sampleBytes, littleEndian) - r[2]) * (maximum / 2) / (r[3] - r[2]);
            double cr = (ReadTiffJpegSample(pixels, i + 2, sampleBytes, littleEndian) - r[4]) * (maximum / 2) / (r[5] - r[4]);
            double red = y + cr * (2 - 2 * c[0]), blue = y + cb * (2 - 2 * c[2]);
            WriteTiffJpegSample(pixels, i, sampleBytes, littleEndian, Clamp(red));
            WriteTiffJpegSample(pixels, i + 1, sampleBytes, littleEndian, Clamp((y - c[0] * red - c[2] * blue) / c[1]));
            WriteTiffJpegSample(pixels, i + 2, sampleBytes, littleEndian, Clamp(blue));
        }
    }
}
