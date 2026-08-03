using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal static class PowerPointBenchmarkVisualValidator {
    internal static int CountPixelsDifferentFrom(
        OfficeRasterImage image,
        OfficeColor background,
        double leftPoints,
        double topPoints,
        double widthPoints,
        double heightPoints) {
        int left = Math.Max(0,
            (int)Math.Floor(leftPoints / 960D * image.Width));
        int top = Math.Max(0,
            (int)Math.Floor(topPoints / 540D * image.Height));
        int right = Math.Min(image.Width,
            (int)Math.Ceiling((leftPoints + widthPoints) / 960D
                * image.Width));
        int bottom = Math.Min(image.Height,
            (int)Math.Ceiling((topPoints + heightPoints) / 540D
                * image.Height));
        int different = 0;
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.A > 0 && Math.Abs(pixel.R - background.R)
                        + Math.Abs(pixel.G - background.G)
                        + Math.Abs(pixel.B - background.B) > 12) {
                    different++;
                }
            }
        }
        return different;
    }
}
