using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Theory]
    [InlineData(false, 48)]
    [InlineData(true, 48)]
    [InlineData(false, 2049)]
    [InlineData(true, 2049)]
    public void ConsecutiveContourFillsKeepIndependentCoverageAfterGrowthAndInvalidGeometry(bool nonZero, int columns) {
        var contours = Enumerable.Range(0, columns).Select(i => (IReadOnlyList<OfficePoint>)new[] {
            new OfficePoint(2D * i + 0.25D, -0.25D), new OfficePoint(2D * i + 1.5D, -0.25D),
            new OfficePoint(2D * i + 1.5D, 3.25D), new OfficePoint(2D * i + 0.25D, 3.25D)
        }).ToArray();
        var image = new OfficeRasterImage(2 * columns, 8, OfficeColor.Transparent);
        var canvas = new OfficeRasterCanvas(image);
        var color = OfficeColor.FromRgba(200, 40, 80, 128);
        // Shared vertex heights must not exhaust the work limit on the final
        // fractional row, even after scratch buffers grow for a wide fill.
        Fill(contours);
        if (columns == 2049) {
            // Distinct subdivisions still exceed the public work limit. The next
            // fill must not reuse scratch state left by that rejected operation.
            var excessive = Enumerable.Range(0, 6000).Select(i =>
                new OfficePoint(i % 2 == 0 ? 0.1D : 0.9D, 4.1D + 0.8D * i / 6000D)).ToArray();
            Assert.Throws<InvalidOperationException>(() => Fill(new[] { excessive }));
        }
        // Rejected geometry must not leave partial contour state in the next fill.
        Fill(new[] { new[] { new OfficePoint(0, 4), new OfficePoint(5, 4), new OfficePoint(double.NaN, 6) } });
        Fill(new[] { new[] { new OfficePoint(0.25D, 5.25D), new OfficePoint(5.75D, 5.25D),
            new OfficePoint(5.75D, 7.75D), new OfficePoint(0.25D, 7.75D) } });

        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                double firstArea = (x % 2 == 0 ? 0.75D : 0.5D)
                    * Math.Max(0D, Math.Min(y + 1D, 3.25D) - y);
                double secondArea = Math.Max(0D, Math.Min(x + 1D, 5.75D) - Math.Max(x, 0.25D))
                    * Math.Max(0D, Math.Min(y + 1D, 7.75D) - Math.Max(y, 5.25D));
                Assert.Equal((byte)Math.Round(128D * (firstArea + secondArea)), image.GetPixel(x, y).A);
            }
        }

        void Fill(IReadOnlyList<IReadOnlyList<OfficePoint>> shapes) {
            if (nonZero) canvas.FillPolygonsNonZero(shapes, color);
            else canvas.FillPolygonsEvenOdd(shapes, color);
        }
    }

    [Theory]
    [InlineData(false, 48)]
    [InlineData(true, 48)]
    [InlineData(false, 2049)]
    [InlineData(true, 2049)]
    [InlineData(false, 5000)]
    [InlineData(true, 5000)]
    public void ManyDisjointContourColumnsRetainCoverageAcrossScratchGrowth(bool nonZero, int columns) {
        var contours = Enumerable.Range(0, columns).Select(i => (IReadOnlyList<OfficePoint>)new[] {
            new OfficePoint(2D * i + 0.25D, -0.25D), new OfficePoint(2D * i + 1.5D, -0.25D),
            new OfficePoint(2D * i + 1.5D, 3.25D), new OfficePoint(2D * i + 0.25D, 3.25D)
        }).ToArray();
        var image = new OfficeRasterImage(2 * columns, 3, OfficeColor.Transparent);
        var canvas = new OfficeRasterCanvas(image);
        var color = OfficeColor.FromRgba(200, 40, 80, 128);
        if (nonZero) canvas.FillPolygonsNonZero(contours, color);
        else canvas.FillPolygonsEvenOdd(contours, color);

        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                double horizontal = x % 2 == 0 ? 0.75D : 0.5D;
                Assert.Equal((byte)Math.Round(128D * horizontal), image.GetPixel(x, y).A);
            }
        }
    }

    [Theory]
    [InlineData(false, 0D)]
    [InlineData(true, 0D)]
    [InlineData(false, 0.0625D)]
    [InlineData(true, 0.0625D)]
    [InlineData(false, 0.999999999999D)]
    [InlineData(true, 0.999999999999D)]
    public void SeparatedContourBandsKeepIndependentHoleAreasAcrossEmptyRows(bool nonZero, double phase) {
        var contours = new List<IReadOnlyList<OfficePoint>>();
        var bands = new[] { -0.5D + phase, 3D + phase, 7.25D + phase };
        foreach (double top in bands) {
            contours.Add(Rectangle(0.25D, top, 5.75D, top + 1.5D));
            contours.Add(Enumerable.Reverse(Rectangle(1.25D, top + 0.25D, 4.25D, top + 1.25D)).ToArray());
        }
        var image = new OfficeRasterImage(7, 11, OfficeColor.Transparent);
        var canvas = new OfficeRasterCanvas(image);
        var color = OfficeColor.FromRgba(200, 40, 80, 128);
        if (nonZero) canvas.FillPolygonsNonZero(contours, color);
        else canvas.FillPolygonsEvenOdd(contours, color);

        // Rectangle intersection areas provide an independent oracle for every pixel,
        // including clipped bands, empty rows and both ends of fractional holes.
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                double area = bands.Sum(top => Area(x, y, 0.25D, top, 5.75D, top + 1.5D)
                    - Area(x, y, 1.25D, top + 0.25D, 4.25D, top + 1.25D));
                Assert.Equal((byte)Math.Round(128D * area), image.GetPixel(x, y).A);
            }
        }

        static OfficePoint[] Rectangle(double left, double top, double right, double bottom) => new[] {
            new OfficePoint(left, top), new OfficePoint(right, top),
            new OfficePoint(right, bottom), new OfficePoint(left, bottom)
        };
        static double Area(int x, int y, double left, double top, double right, double bottom) =>
            Math.Max(0D, Math.Min(x + 1D, right) - Math.Max(x, left))
            * Math.Max(0D, Math.Min(y + 1D, bottom) - Math.Max(y, top));
    }

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
