using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FractionalContourRowsRetainHoleCoverageAndCompositeAlphaOnce(bool nonZero) {
        var outer = new[] {
            new OfficePoint(0.25, 0.25), new OfficePoint(4.75, 0.25),
            new OfficePoint(4.75, 2.75), new OfficePoint(0.25, 2.75)
        };
        // Opposite winding makes the hole equivalent under both fill rules.
        var hole = new[] {
            new OfficePoint(1.25, 0.75), new OfficePoint(1.25, 2.25),
            new OfficePoint(3.75, 2.25), new OfficePoint(3.75, 0.75)
        };
        var image = new OfficeRasterImage(6, 4, OfficeColor.Transparent);
        var canvas = new OfficeRasterCanvas(image);
        var color = OfficeColor.FromRgba(200, 40, 80, 128);
        if (nonZero) canvas.FillPolygonsNonZero(new[] { outer, hole }, color);
        else canvas.FillPolygonsEvenOdd(new[] { outer, hole }, color);

        // Independent rectangle intersection areas: .75*.75 and
        // 1*.75 - .75*.25 both cover .5625 of a pixel; .25 covers the hole edge.
        Assert.Equal(72, image.GetPixel(0, 0).A);
        Assert.Equal(72, image.GetPixel(1, 0).A);
        Assert.Equal(32, image.GetPixel(1, 1).A);
        Assert.Equal(0, image.GetPixel(2, 1).A);
        Assert.Equal(72, image.GetPixel(4, 2).A);
        Assert.Equal(0, image.GetPixel(5, 2).A);
        Assert.Equal(0, image.GetPixel(2, 3).A);
    }
}
