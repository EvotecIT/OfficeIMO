using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingDisjointContourGroupsTests {
    private static readonly OfficeColor HalfInk = new(20, 40, 60, 128);

    [Fact]
    public void ContoursSharingAPixelAreFilledTogether() {
        // The second glyph starts inside the first one's last pixel column; filling them separately
        // would blend that column twice with the translucent color.
        var contours = new List<IReadOnlyList<OfficePoint>> { Glyph(2.3, 10.6), Glyph(10.2, 18.7) };
        byte[] single = Fill(canvas => canvas.FillContourPaint(contours, OfficeFillRule.NonZero, OfficeRasterCanvas.SolidPaint(HalfInk)));
        byte[] grouped = Fill(canvas => canvas.FillDisjointContourGroups(contours, OfficeFillRule.NonZero, OfficeRasterCanvas.SolidPaint(HalfInk)));
        Assert.Equal(single, grouped);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SeparateGlyphsDifferOnlyInEdgeAntialiasing(bool syntheticBold) {
        var contours = new List<IReadOnlyList<OfficePoint>> { Glyph(2.3, 9.6), Glyph(11.4, 18.7), Glyph(20.1, 27.2) };
        List<IReadOnlyList<OfficePoint>>? shifted = syntheticBold
            ? contours.Select(contour => (IReadOnlyList<OfficePoint>)contour.Select(point => new OfficePoint(point.X + 0.4, point.Y)).ToList()).ToList()
            : null;
        byte[] single = Fill(canvas => canvas.FillContourPaint(contours, OfficeFillRule.NonZero, OfficeRasterCanvas.SolidPaint(HalfInk), shifted));
        byte[] grouped = Fill(canvas => canvas.FillDisjointContourGroups(contours, OfficeFillRule.NonZero, OfficeRasterCanvas.SolidPaint(HalfInk), shifted));
        // Sub-scanlines now split only at the group's own vertices, so edge pixels may move by a few levels.
        for (int index = 3; index < single.Length; index += 4) Assert.InRange(Math.Abs(single[index] - grouped[index]), 0, 8);
    }

    private static byte[] Fill(Action<OfficeRasterCanvas> paint) {
        var image = new OfficeRasterImage(32, 20);
        paint(new OfficeRasterCanvas(image));
        return image.GetPixels();
    }

    // A notched wedge; its vertex heights move with its position, as different glyphs' heights differ.
    private static List<OfficePoint> Glyph(double left, double right) {
        double middle = (left + right) / 2, lift = left * 0.037;
        return new List<OfficePoint> {
            new(left, 17.3 - lift), new(middle - 0.7, 2.6 + lift), new(middle + 0.9, 3.1 + lift), new(right, 16.4 - lift), new(middle, 12.9 - lift)
        };
    }
}
