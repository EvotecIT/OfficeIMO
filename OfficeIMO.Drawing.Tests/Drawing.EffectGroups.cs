using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingEffectGroup_CompositesOpacityOnceAfterTransform() {
        var inner = new OfficeDrawing(20D, 20D);
        OfficeShape first = OfficeShape.Rectangle(10D, 10D);
        first.FillColor = OfficeColor.Red;
        first.StrokeWidth = 0D;
        OfficeShape second = first.Clone();
        inner.AddShape(first, 0D, 0D);
        inner.AddShape(second, 0D, 0D);

        var drawing = new OfficeDrawing(40D, 30D);
        drawing.AddEffectDrawing(inner, OfficeTransform.Translate(10D, 5D), 0.5D);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeColor painted = raster.GetPixel(12, 7);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(2, 2));
        Assert.Equal((byte)255, painted.R);
        Assert.Equal((byte)0, painted.G);
        Assert.Equal((byte)0, painted.B);
        Assert.InRange(painted.A, (byte)127, (byte)128);
        Assert.Contains("opacity=\"0.5\"", svg, StringComparison.Ordinal);
        Assert.Contains("matrix(1 0 0 1 10 5)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingEffectGroup_AppliesArbitraryAffineScale() {
        var inner = new OfficeDrawing(10D, 10D);
        OfficeShape shape = OfficeShape.Rectangle(10D, 10D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        inner.AddShape(shape, 0D, 0D);

        var drawing = new OfficeDrawing(40D, 30D);
        OfficeTransform transform = OfficeTransform.Scale(2D, 1.5D).Then(OfficeTransform.Translate(5D, 4D));
        drawing.AddEffectDrawing(inner, transform);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(6, 5));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(23, 17));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(26, 17));
    }

    [Fact]
    public void OfficeDrawingEffectGroup_AppliesAffineRotation() {
        var inner = new OfficeDrawing(10D, 20D);
        OfficeShape shape = OfficeShape.Rectangle(10D, 20D);
        shape.FillColor = OfficeColor.Blue;
        shape.StrokeWidth = 0D;
        inner.AddShape(shape, 0D, 0D);

        var drawing = new OfficeDrawing(35D, 30D);
        OfficeTransform transform = OfficeTransform.RotateDegrees(90D).Then(OfficeTransform.Translate(25D, 5D));
        drawing.AddEffectDrawing(inner, transform);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(6, 6));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(23, 13));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(10, 20));
        Assert.Contains("matrix(0 1 -1 0 25 5)", svg, StringComparison.Ordinal);
    }
}
