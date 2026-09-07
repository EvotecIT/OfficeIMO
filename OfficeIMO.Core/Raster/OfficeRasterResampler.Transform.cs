using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeRasterResampler {
    /// <summary>Resamples an affine-transformed image into a bounded canvas, filling uncovered pixels with the background.</summary>
    /// <remarks>Coordinates describe pixel edges. Source pixels are never modified. Sampling uses premultiplied alpha.</remarks>
    public static OfficeRasterImage Transform(OfficeRasterImage source, OfficeTransform sourceToDestination,
        int width, int height, OfficeColor? background = null, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        OfficeRasterGuards.EnsureOutputPixels(width, height, "Transformed raster exceeds the managed image limit.");
        EnsureSimpleWorkingSet(source, width, height, retainedManagedBytes: 0L);
        OfficeTransform inverse = sourceToDestination.Invert();
        OfficeColor fill = background ?? OfficeColor.White;
        cancellationToken.ThrowIfCancellationRequested();
        var result = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if ((x & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                OfficePoint point = inverse.TransformPoint(new OfficePoint(x + 0.5D, y + 0.5D));
                double sx = point.X - 0.5D, sy = point.Y - 0.5D;
                // Avoid narrowing enormous inverse coordinates before rejecting pixels outside the source.
                if (sx < -1D || sy < -1D || sx > source.Width || sy > source.Height) {
                    result.SetPixel(x, y, fill);
                    continue;
                }
                int left = (int)Math.Floor(sx), top = (int)Math.Floor(sy);
                double fx = sx - left, fy = sy - top;
                double r = 0D, g = 0D, b = 0D, a = 0D;
                Add(left, top, (1D - fx) * (1D - fy));
                Add(left + 1, top, fx * (1D - fy));
                Add(left, top + 1, (1D - fx) * fy);
                Add(left + 1, top + 1, fx * fy);
                result.SetPixel(x, y, a <= 0D ? OfficeColor.FromRgba(0, 0, 0, 0) :
                    OfficeColor.FromRgba(Channel(r / a), Channel(g / a), Channel(b / a), Channel(a * 255D)));

                void Add(int px, int py, double weight) {
                    OfficeColor color = px < 0 || py < 0 || px >= source.Width || py >= source.Height ? fill : source.GetPixel(px, py);
                    double alpha = color.A / 255D * weight;
                    r += color.R * alpha; g += color.G * alpha; b += color.B * alpha; a += alpha;
                }
            }
        }
        return result;
    }

    private static byte Channel(double value) => (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));
}
