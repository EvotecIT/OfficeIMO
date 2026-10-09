using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    /// <summary>Multiplies encoded RGB brightness; one leaves the color unchanged.</summary>
    public static OfficeRasterImage Brightness(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount));
        return Map(source, (c, _, _) => Rgb(c.R * amount, c.G * amount, c.B * amount, c.A), cancellationToken);
    }

    /// <summary>Multiplies contrast around the encoded-channel midpoint; one is the identity.</summary>
    public static OfficeRasterImage Contrast(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount));
        return Map(source, (c, _, _) => Rgb((c.R - 127.5D) * amount + 127.5D, (c.G - 127.5D) * amount + 127.5D, (c.B - 127.5D) * amount + 127.5D, c.A), cancellationToken);
    }

    /// <summary>Scales color distance from BT.709 luminance; zero produces grayscale and one is the identity.</summary>
    public static OfficeRasterImage Saturate(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount));
        return Map(source, (c, _, _) => {
            double gray = Luminance(c);
            return Rgb(gray + (c.R - gray) * amount, gray + (c.G - gray) * amount, gray + (c.B - gray) * amount, c.A);
        }, cancellationToken);
    }

    /// <summary>Scales HSL lightness while retaining hue, saturation, and alpha; one is the identity.</summary>
    public static OfficeRasterImage Lightness(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount));
        return Map(source, (c, _, _) => {
            ToHsl(c, out double h, out double s, out double l);
            return FromHsl(h, s, Math.Min(1D, l * amount), c.A);
        }, cancellationToken);
    }

    /// <summary>Multiplies alpha by an amount between zero and one, retaining straight RGB.</summary>
    public static OfficeRasterImage Opacity(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount), maximum: 1D);
        return Map(source, (c, _, _) => OfficeColor.FromRgba(c.R, c.G, c.B, Channel(c.A * amount)), cancellationToken);
    }

    /// <summary>Rotates HSL hue by the supplied number of degrees, retaining lightness and alpha.</summary>
    public static OfficeRasterImage Hue(OfficeRasterImage source, double degrees, CancellationToken cancellationToken = default) {
        ValidateAmount(degrees, nameof(degrees), -1_000_000D, 1_000_000D);
        return Map(source, (c, _, _) => {
            ToHsl(c, out double h, out double s, out double l);
            h = ((h + degrees / 360D) % 1D + 1D) % 1D;
            return FromHsl(h, s, l, c.A);
        }, cancellationToken);
    }

    /// <summary>Inverts encoded RGB while retaining alpha.</summary>
    public static OfficeRasterImage Invert(OfficeRasterImage source, CancellationToken cancellationToken = default) =>
        Map(source, (c, _, _) => Rgb(255 - c.R, 255 - c.G, 255 - c.B, c.A), cancellationToken);

    /// <summary>Converts RGB to the selected encoded-channel luminance while retaining alpha.</summary>
    public static OfficeRasterImage Grayscale(OfficeRasterImage source, OfficeRasterGrayscaleMode mode = OfficeRasterGrayscaleMode.Bt709, CancellationToken cancellationToken = default) {
        if (mode != OfficeRasterGrayscaleMode.Bt709 && mode != OfficeRasterGrayscaleMode.Bt601) throw new ArgumentOutOfRangeException(nameof(mode));
        return Map(source, (c, _, _) => {
            byte gray = Channel(mode == OfficeRasterGrayscaleMode.Bt709 ? Luminance(c) : c.R * .299D + c.G * .587D + c.B * .114D);
            return OfficeColor.FromRgba(gray, gray, gray, c.A);
        }, cancellationToken);
    }

    /// <summary>Produces black or white using a normalized BT.709 threshold, retaining alpha.</summary>
    public static OfficeRasterImage Threshold(OfficeRasterImage source, double threshold = .5D, CancellationToken cancellationToken = default) {
        ValidateAmount(threshold, nameof(threshold), maximum: 1D);
        return Map(source, (c, _, _) => {
            byte channel = Luminance(c) >= threshold * 255D ? (byte)255 : (byte)0;
            return OfficeColor.FromRgba(channel, channel, channel, c.A);
        }, cancellationToken);
    }

    /// <summary>Applies a caller-owned normalized RGBA affine matrix, including its explicit alpha row.</summary>
    public static OfficeRasterImage ColorMatrix(OfficeRasterImage source, OfficeRasterColorMatrix matrix, CancellationToken cancellationToken = default) {
        if (matrix == null) throw new ArgumentNullException(nameof(matrix));
        return Map(source, (c, _, _) => {
            double r = c.R / 255D, g = c.G / 255D, b = c.B / 255D, a = c.A / 255D;
            return OfficeColor.FromRgba(Channel(matrix.Apply(0, r, g, b, a) * 255D), Channel(matrix.Apply(1, r, g, b, a) * 255D),
                Channel(matrix.Apply(2, r, g, b, a) * 255D), Channel(matrix.Apply(3, r, g, b, a) * 255D));
        }, cancellationToken);
    }

    /// <summary>Interpolates between the original color and a warm sepia transform, retaining alpha.</summary>
    public static OfficeRasterImage Sepia(OfficeRasterImage source, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount), maximum: 1D);
        return Map(source, (c, _, _) => Rgb(
            c.R * (1D - amount) + amount * (.393D * c.R + .769D * c.G + .189D * c.B),
            c.G * (1D - amount) + amount * (.349D * c.R + .686D * c.G + .168D * c.B),
            c.B * (1D - amount) + amount * (.272D * c.R + .534D * c.G + .131D * c.B), c.A), cancellationToken);
    }

    private static void ToHsl(OfficeColor c, out double hue, out double saturation, out double lightness) {
        double r = c.R / 255D, g = c.G / 255D, b = c.B / 255D;
        double maximum = Math.Max(r, Math.Max(g, b)), minimum = Math.Min(r, Math.Min(g, b));
        double delta = maximum - minimum;
        lightness = (maximum + minimum) / 2D;
        if (delta < 1E-12D) { hue = saturation = 0D; return; }
        saturation = delta / (1D - Math.Abs(2D * lightness - 1D));
        hue = maximum == r ? ((g - b) / delta + (g < b ? 6D : 0D)) / 6D : maximum == g ? ((b - r) / delta + 2D) / 6D : ((r - g) / delta + 4D) / 6D;
    }

    private static OfficeColor FromHsl(double h, double s, double l, byte alpha) {
        double chroma = (1D - Math.Abs(2D * l - 1D)) * s;
        double x = chroma * (1D - Math.Abs((h * 6D) % 2D - 1D)), m = l - chroma / 2D;
        int sector = (int)Math.Floor(h * 6D) % 6;
        return sector switch {
            0 => Rgb((chroma + m) * 255D, (x + m) * 255D, m * 255D, alpha),
            1 => Rgb((x + m) * 255D, (chroma + m) * 255D, m * 255D, alpha),
            2 => Rgb(m * 255D, (chroma + m) * 255D, (x + m) * 255D, alpha),
            3 => Rgb(m * 255D, (x + m) * 255D, (chroma + m) * 255D, alpha),
            4 => Rgb((x + m) * 255D, m * 255D, (chroma + m) * 255D, alpha),
            _ => Rgb((chroma + m) * 255D, m * 255D, (x + m) * 255D, alpha)
        };
    }
}
