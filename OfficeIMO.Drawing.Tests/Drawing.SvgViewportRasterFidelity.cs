using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgViewportRasterFidelityTests {
    private const string Bars = "<svg xmlns='http://www.w3.org/2000/svg' width='60' height='24' viewBox='0 0 20 8'>"
        + "<rect width='20' height='8' fill='white'/>"
        + "<g fill='black'><rect x='2' width='2' height='6'/><rect x='5' width='1' height='6'/>"
        + "<rect x='7' width='1' height='6'/><rect x='10' width='3' height='6'/></g></svg>";

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ScaledSvgViewportPreservesIntegerAlignedBarEdges(int outputScale) {
        OfficeDrawing drawing = ReadBars();

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, outputScale);

        AssertBars(image, outputScale);
    }

    [Fact]
    public void PointConversionAndDpiScalingPreserveOriginalSvgBarEdges() {
        OfficeDrawing pixels = ReadBars();
        var points = new OfficeDrawing(pixels.Width * 0.75D, pixels.Height * 0.75D);
        points.AddEffectDrawing(pixels, OfficeTransform.Scale(0.75D, 0.75D));

        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(points, 96D / 72D);

        AssertBars(image, 1);
    }

    [Theory]
    [InlineData(0.5D)]
    [InlineData(1.25D)]
    public void FractionalEffectScaleMatchesDirectVectorCoverage(double groupScale) {
        OfficeDrawing svg = ReadBars();
        var transformed = new OfficeDrawing(svg.Width * groupScale, svg.Height * groupScale);
        transformed.AddEffectDrawing(svg, OfficeTransform.Scale(groupScale, groupScale));
        var reference = new OfficeDrawing(20D, 8D);
        AddRectangle(reference, 0D, 20D, 8D, OfficeColor.White);
        foreach ((double x, double width) in new[] { (2D, 2D), (5D, 1D), (7D, 1D), (10D, 3D) }) {
            AddRectangle(reference, x, width, 6D, OfficeColor.Black);
        }

        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(transformed);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(reference, 3D * groupScale);

        Assert.Equal(expected.Width, actual.Width);
        Assert.Equal(expected.Height, actual.Height);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
    }

    [Fact]
    public void MagnifiedEffectLayerHonorsIntermediatePixelBudget() {
        var inner = new OfficeDrawing(100D, 100D);
        var rectangle = OfficeShape.Rectangle(100D, 100D);
        rectangle.FillColor = OfficeColor.Black;
        rectangle.StrokeColor = null;
        inner.AddShape(rectangle, 0D, 0D);
        var drawing = new OfficeDrawing(100D, 100D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Scale(8D, 8D));

        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10_000 }));
    }

    private static OfficeDrawing ReadBars() {
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(Bars), out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        return drawing!;
    }

    private static void AddRectangle(OfficeDrawing drawing, double x, double width, double height, OfficeColor color) {
        OfficeShape rectangle = OfficeShape.Rectangle(width, height);
        rectangle.FillColor = color;
        rectangle.StrokeColor = null;
        drawing.AddShape(rectangle, x, 0D);
    }

    private static void AssertBars(OfficeRasterImage image, int outputScale) {
        Assert.Equal(60 * outputScale, image.Width);
        Assert.Equal(24 * outputScale, image.Height);
        for (int x = 0; x < image.Width; x++) {
            int module = x / (3 * outputScale);
            bool dark = module is 2 or 3 or 5 or 7 or 10 or 11 or 12;
            Assert.Equal(dark ? OfficeColor.Black : OfficeColor.White, image.GetPixel(x, 6 * outputScale));
        }
    }
}
