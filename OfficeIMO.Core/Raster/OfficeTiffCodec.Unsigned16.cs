using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    private static void ConvertUnsigned16Pixel(
        byte[] source, int offset, bool littleEndian, int photometric,
        int alphaIndex, int alphaKind, double[]? colorComponents,
        out byte red, out byte green, out byte blue, out byte alpha) {
        int alphaSample = alphaIndex >= 0
            ? ReadUInt16(source, offset + alphaIndex * 2, littleEndian)
            : ushort.MaxValue;
        alpha = ColorMapByte(alphaSample);

        // Keep sample and associated-alpha precision until the final RGBA projection
        // or ICC transform. Rounding alpha to eight bits first can erase low alpha.
        double Component(int channel) {
            if (alphaKind == 1 && alphaSample == 0) return 0D;
            double denominator = alphaKind == 1 ? alphaSample : ushort.MaxValue;
            return Math.Min(1D, ReadUInt16(source, offset + channel * 2, littleEndian) / denominator);
        }

        if (photometric == 0 || photometric == 1) {
            double gray = Component(0);
            if (photometric == 0) gray = 1D - gray;
            if (colorComponents != null) colorComponents[0] = gray;
            red = green = blue = QuantizeUnsigned16Component(gray);
        } else if (photometric == 2) {
            double r = Component(0), g = Component(1), b = Component(2);
            if (colorComponents != null) {
                colorComponents[0] = r; colorComponents[1] = g; colorComponents[2] = b;
            }
            red = QuantizeUnsigned16Component(r);
            green = QuantizeUnsigned16Component(g);
            blue = QuantizeUnsigned16Component(b);
        } else {
            double c = Component(0), m = Component(1), y = Component(2), k = Component(3);
            if (colorComponents != null) {
                colorComponents[0] = c; colorComponents[1] = m;
                colorComponents[2] = y; colorComponents[3] = k;
            }
            // Preserve the ordinary decoder's device-CMYK approximation. XPS
            // requires a usable CMYK ICC profile before accepting these pixels.
            int black = QuantizeUnsigned16Component(k);
            red = (byte)(255 - Math.Min(255, QuantizeUnsigned16Component(c) + black));
            green = (byte)(255 - Math.Min(255, QuantizeUnsigned16Component(m) + black));
            blue = (byte)(255 - Math.Min(255, QuantizeUnsigned16Component(y) + black));
        }
    }

    private static byte QuantizeUnsigned16Component(double value) =>
        (byte)Math.Floor(value * 255D + 0.5D);
}
