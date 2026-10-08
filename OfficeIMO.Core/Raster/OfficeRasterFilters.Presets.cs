using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterFilters {
    /// <summary>Blends the outer elliptical image region toward a color, retaining source alpha; the center remains unchanged.</summary>
    public static OfficeRasterImage Vignette(OfficeRasterImage source, OfficeColor? color = null, double amount = 1D, CancellationToken cancellationToken = default) {
        ValidateAmount(amount, nameof(amount), maximum: 1D);
        if (source == null) throw new ArgumentNullException(nameof(source));
        OfficeColor tint = color ?? OfficeColor.Black;
        return Map(source, (c, x, y) => VignettePixel(c, tint, amount, x, y, source.Width, source.Height), cancellationToken);
    }

    /// <summary>Applies a warm, saturated, contrast-enhanced film-style color preset while retaining alpha.</summary>
    public static OfficeRasterImage Kodachrome(OfficeRasterImage source, CancellationToken cancellationToken = default) =>
        Map(source, (c, _, _) => Rgb((c.R - 127.5D) * 1.2D + 135D, (c.G - 127.5D) * 1.1D + 130D, (c.B - 127.5D) * 1.1D + 118D, c.A), cancellationToken);

    /// <summary>Applies a saturated, green-biased film-style preset with dark elliptical edges, retaining alpha.</summary>
    public static OfficeRasterImage Lomograph(OfficeRasterImage source, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return Map(source, (c, x, y) => {
            double gray = Luminance(c);
            OfficeColor film = Rgb(gray + (c.R - gray) * 1.3D, gray + (c.G - gray) * 1.3D + 10D, gray + (c.B - gray) * 1.3D - 8D, c.A);
            return VignettePixel(film, OfficeColor.Black, .6D, x, y, source.Width, source.Height);
        }, cancellationToken);
    }

    /// <summary>Applies a faded warm instant-film-style preset with gentle edge tint, retaining alpha.</summary>
    public static OfficeRasterImage Polaroid(OfficeRasterImage source, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return Map(source, (c, x, y) => {
            OfficeColor film = Rgb(c.R * .85D + 30D, c.G * .82D + 25D, c.B * .78D + 20D, c.A);
            return VignettePixel(film, OfficeColor.FromRgb(65, 45, 30), .35D, x, y, source.Width, source.Height);
        }, cancellationToken);
    }

    private static OfficeColor VignettePixel(OfficeColor pixel, OfficeColor tint, double amount, int x, int y, int width, int height) {
        double dx = (x + .5D - width / 2D) / Math.Max(.5D, width / 2D);
        double dy = (y + .5D - height / 2D) / Math.Max(.5D, height / 2D);
        double t = Math.Max(0D, Math.Min(1D, (Math.Sqrt(dx * dx + dy * dy) - .35D) / .95D));
        double blend = t * t * (3D - 2D * t) * amount * tint.A / 255D;
        return Rgb(pixel.R * (1D - blend) + tint.R * blend, pixel.G * (1D - blend) + tint.G * blend, pixel.B * (1D - blend) + tint.B * blend, pixel.A);
    }
}
