using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeDrawingRasterRenderer_UnpaintedLinesDoNotAcquireFillOrBlackStroke(bool transformed) {
        var drawing = new OfficeDrawing(100, 80);
        var line = OfficeShape.Line(0, 0, 50, 20); line.FillColor = OfficeColor.Red; line.StrokeColor = null; line.StrokeWidth = 6;
        if (transformed) line.Transform = OfficeTransform.Translate(5, 5);
        drawing.AddShape(line, 20, 20);
        Assert.Contains("stroke=\"none\"", OfficeDrawingSvgExporter.ToSvg(drawing));
        var image = OfficeDrawingRasterRenderer.Render(drawing);
        for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++) Assert.Equal(0, image.GetPixel(x, y).A);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeDrawingRasterRenderer_GradientOnlyLinesRemainPaintedWithoutSolidColor(bool transformed) {
        var drawing = new OfficeDrawing(100, 80);
        var line = OfficeShape.Line(0, 0, 50, 20); line.StrokeWidth = 6;
        line.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        if (transformed) line.Transform = OfficeTransform.Translate(5, 5);
        drawing.AddShape(line, 20, 20);
        var image = OfficeDrawingRasterRenderer.Render(drawing);
        int shift = transformed ? 5 : 0; var first = image.GetPixel(25 + shift, 22 + shift); var last = image.GetPixel(65 + shift, 38 + shift);
        Assert.True(first.A > 0 && first.R > first.B); Assert.True(last.A > 0 && last.B > last.R);
    }
}
