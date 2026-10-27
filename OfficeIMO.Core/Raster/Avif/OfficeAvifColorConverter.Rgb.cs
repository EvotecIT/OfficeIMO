using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeAvifColorConverter {
    /// <summary>Converts OS-decoded source RGB samples to sRGB once, after the codec has reconstructed YUV.</summary>
    internal static void ConvertRgbToSrgb(OfficeRasterImage image, ushort primaries, ushort transfer, CancellationToken token) {
        if (primaries is 1 or 2 && transfer is 2 or 13) return;
        if (primaries is not (1 or 2 or 9 or 12) || transfer is not (1 or 2 or 6 or 13 or 14 or 15))
            throw new NotSupportedException("HEIF nclx primaries or transfer function is outside the supported SDR subset.");
        byte[] rgba = image.PixelBuffer;
        for (int i = 0; i < rgba.Length; i += 4) {
            if ((i & 16383) == 0) token.ThrowIfCancellationRequested();
            double r = Linear(rgba[i] / 255D, transfer), g = Linear(rgba[i + 1] / 255D, transfer), b = Linear(rgba[i + 2] / 255D, transfer);
            double sr = r, sg = g, sb = b;
            if (primaries == 12) {
                sr = 1.2249401763 * r - .2249401763 * g;
                sg = -.0420569547 * r + 1.0420569547 * g;
                sb = -.0196375546 * r - .0786360456 * g + 1.0982736002 * b;
            } else if (primaries == 9) {
                sr = 1.6604910021 * r - .5876411388 * g - .0728498633 * b;
                sg = -.1245504745 * r + 1.1328998971 * g - .0083494226 * b;
                sb = -.0181507634 * r - .1005788980 * g + 1.1187296614 * b;
            }
            rgba[i] = Srgb(sr); rgba[i + 1] = Srgb(sg); rgba[i + 2] = Srgb(sb);
        }
    }

    private static double Linear(double v, ushort transfer) => transfer is 2 or 13
        ? (v <= .04045 ? v / 12.92 : Math.Pow((v + .055) / 1.055, 2.4))
        : (v < .081 ? v / 4.5 : Math.Pow((v + .099) / 1.099, 1 / .45));

    private static byte Srgb(double linear) {
        double v = linear <= .0031308 ? 12.92 * linear : 1.055 * Math.Pow(linear, 1 / 2.4) - .055;
        return (byte)Math.Max(0, Math.Min(255, Math.Floor(v * 255 + .5)));
    }
}
