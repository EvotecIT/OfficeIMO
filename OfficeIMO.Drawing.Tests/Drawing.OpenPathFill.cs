using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingOpenPathFillTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RasterFillClosesOpenContourWithoutAddingClosingStroke(bool transformed) {
        var shape = OfficeShape.Path(40, 40, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(40, 0), OfficePathCommand.LineTo(40, 40));
        shape.FillColor = OfficeColor.Red;
        shape.StrokeColor = OfficeColor.Blue;
        shape.StrokeWidth = 2;
        if (transformed) shape.Transform = OfficeTransform.Translate(2, 2);
        var drawing = new OfficeDrawing(60, 60).AddShape(shape, 10, 10);
        var raster = OfficeDrawingRasterRenderer.Render(drawing, 1, OfficeColor.White);
        int offset = transformed ? 2 : 0;
        Assert.Equal(OfficeColor.Red, raster.GetPixel(40 + offset, 20 + offset));
        Assert.Equal(OfficeColor.White, raster.GetPixel(20 + offset, 40 + offset));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(30 + offset, 29 + offset));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(49 + offset, 30 + offset));
    }
}
