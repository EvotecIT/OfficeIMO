using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Tests {
    internal static class RasterColorTestSupport {
        // Measure painted area, including antialiasing, instead of requiring a thin
        // stroke or glyph to contain pixels identical to its foreground color.
        internal static double MeasureCoverage(OfficeRasterImage image, OfficeColor foreground, OfficeColor background) {
            double coverage = 0D;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    coverage += GetCoverage(image.GetPixel(x, y), foreground, background);
                }
            }
            return coverage;
        }

        internal static double GetCoverage(OfficeColor actual, OfficeColor foreground, OfficeColor background) {
            if (actual.A < 248) return 0D;
            double red = foreground.R - background.R;
            double green = foreground.G - background.G;
            double blue = foreground.B - background.B;
            double lengthSquared = red * red + green * green + blue * blue;
            if (lengthSquared == 0D) return 0D;
            double coverage = ((actual.R - background.R) * red +
                (actual.G - background.G) * green + (actual.B - background.B) * blue) / lengthSquared;
            if (coverage <= 0D || coverage > 1.05D) return 0D;
            coverage = Math.Min(1D, coverage);
            return Math.Abs(actual.R - (background.R + coverage * red)) <= 8D &&
                Math.Abs(actual.G - (background.G + coverage * green)) <= 8D &&
                Math.Abs(actual.B - (background.B + coverage * blue)) <= 8D ? coverage : 0D;
        }

    }
}
