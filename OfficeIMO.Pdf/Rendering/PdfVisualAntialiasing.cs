using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Edge-shift detection used by <see cref="PdfVisualComparisonOptions.IgnoreAntialiasing"/>. When an edge moves by less
/// than a pixel, each changed pixel takes a colour between its neighbours: it stays inside the colour range of the other
/// image's 3x3 neighbourhood. A changed glyph puts ink where the other image's whole neighbourhood is paper, or removes it,
/// so at least one direction falls outside that range.
/// </summary>
internal static class PdfVisualAntialiasing {
    public static bool IsEdgeShift(OfficeRasterImage expected, OfficeRasterImage actual, int x, int y, int tolerance) =>
        IsWithinNeighbourhood(actual.GetPixel(x, y), expected, x, y, tolerance)
        && IsWithinNeighbourhood(expected.GetPixel(x, y), actual, x, y, tolerance);

    private static bool IsWithinNeighbourhood(OfficeColor colour, OfficeRasterImage image, int x, int y, int tolerance) {
        int minR = 255, minG = 255, minB = 255, minA = 255, maxR = 0, maxG = 0, maxB = 0, maxA = 0;
        for (int ny = Math.Max(y - 1, 0); ny <= Math.Min(y + 1, image.Height - 1); ny++) {
            for (int nx = Math.Max(x - 1, 0); nx <= Math.Min(x + 1, image.Width - 1); nx++) {
                OfficeColor neighbour = image.GetPixel(nx, ny);
                minR = Math.Min(minR, neighbour.R); maxR = Math.Max(maxR, neighbour.R);
                minG = Math.Min(minG, neighbour.G); maxG = Math.Max(maxG, neighbour.G);
                minB = Math.Min(minB, neighbour.B); maxB = Math.Max(maxB, neighbour.B);
                minA = Math.Min(minA, neighbour.A); maxA = Math.Max(maxA, neighbour.A);
            }
        }
        return Within(colour.R, minR, maxR, tolerance) && Within(colour.G, minG, maxG, tolerance)
            && Within(colour.B, minB, maxB, tolerance) && Within(colour.A, minA, maxA, tolerance);
    }

    private static bool Within(int value, int minimum, int maximum, int tolerance) =>
        value >= minimum - tolerance && value <= maximum + tolerance;
}
