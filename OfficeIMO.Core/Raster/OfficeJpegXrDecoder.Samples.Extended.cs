using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // T.832 9.10.5-8 and A.7.18: signed samples use s2.13 / s7.24;
    // half/single samples retain their IEEE value until color conversion.
    internal static double ScaleExtendedSample(long value, PlaneHeader plane, int bitDepth) {
        if (plane.Scaled) value = (value + 3) >> 3;
        if (bitDepth == 3 || bitDepth == 6) {
            if (bitDepth == 3) {
                double shifted = value * Math.Pow(2D, plane.ShiftBits);
                return Math.Max(short.MinValue, Math.Min(short.MaxValue, shifted)) / 8192D;
            }
            // BD32S has no clipping stage. A.7.18 packs the low 32 bits;
            // lossy reconstruction can cross a signed endpoint before packing.
            uint packed = plane.ShiftBits >= 32 ? 0U : unchecked((uint)value) << plane.ShiftBits;
            return unchecked((int)packed) / 16777216D;
        }
        long magnitude = Math.Abs(value);
        double sample;
        if (bitDepth == 4) {
            int packed = (int)Math.Min(magnitude, 32767), exponent = packed >> 10, mantissa = packed & 1023;
            if (exponent == 31) throw new FormatException("JPEG-XR non-finite half sample cannot form an SDR pixel.");
            sample = exponent == 0 ? mantissa * Math.Pow(2D, -24) : (1024 + mantissa) * Math.Pow(2D, exponent - 25);
        } else if (bitDepth == 7) {
            long unit = 1L << plane.MantissaBits;
            long exponent = magnitude >> plane.MantissaBits, mantissa = (magnitude & (unit - 1)) | unit;
            if (exponent == 0) { mantissa ^= unit; exponent = 1; }
            exponent += 127 - plane.ExponentBias;
            while (mantissa < unit && exponent > 1 && mantissa > 0) { exponent--; mantissa <<= 1; }
            if (mantissa < unit) exponent = 0;
            else mantissa ^= unit;
            if (exponent < 0 || exponent >= 255)
                throw new FormatException("JPEG-XR floating sample exceeds finite IEEE single precision.");
            mantissa <<= 23 - plane.MantissaBits;
            sample = exponent == 0 ? mantissa * Math.Pow(2D, -149) : (8388608 + mantissa) * Math.Pow(2D, exponent - 150);
        } else throw new FormatException("JPEG-XR extended sample type is unsupported.");
        return value < 0 ? -sample : sample;
    }

    private static void WriteExtendedRgba(byte[] output, int target, double red, double green, double blue, double alpha,
            bool premultiplied, OfficeIccColorProfile? profile, double[]? channels) {
        if (premultiplied) {
            if (alpha <= 0D) red = green = blue = 0D;
            else { red /= alpha; green /= alpha; blue /= alpha; }
        }
        OfficeColor color;
        if (profile != null) {
            channels![0] = red;
            if (channels.Length == 3) { channels[1] = green; channels[2] = blue; }
            if (!profile.TryConvert(channels, OfficeIccRenderingIntent.RelativeColorimetric, out color))
                throw new FormatException("JPEG-XR extended-sample ICC conversion failed.");
        } else color = OfficeColorSpaceConverter.FromLinearSrgb(red, green, blue);
        output[target] = color.R; output[target + 1] = color.G; output[target + 2] = color.B;
        output[target + 3] = (byte)Math.Round(Math.Max(0D, Math.Min(1D, alpha)) * 255D);
    }
}
