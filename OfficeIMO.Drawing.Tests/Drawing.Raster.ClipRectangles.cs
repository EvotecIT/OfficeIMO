using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingRasterClipRectangleTests {
    [Theory]
    [InlineData(0D, 1D, 1)]
    [InlineData(0.25D, 1D, 1)]
    [InlineData(0.5D, 1D, 1)]
    [InlineData(0.499999999999D, 1D, 3)]
    [InlineData(0.500000000001D, 1D, 3)]
    [InlineData(0.25D, 1.5D, 3)]
    public void RectangularGroupClipPreservesPathPixelCentreCoverage(double offset, double scale, int depth) {
        byte[] rectangle = RenderGroups(false, offset, scale, depth);
        byte[] path = RenderGroups(true, offset, scale, depth);

        Assert.Equal(path, rectangle);
        Assert.Contains(rectangle.Where((_, i) => i % 4 == 3), alpha => alpha == 255);
        Assert.Contains(rectangle.Where((_, i) => i % 4 == 3), alpha => alpha == 0);
    }

    [Theory]
    [InlineData(0D, false)]
    [InlineData(90D, false)]
    [InlineData(180D, false)]
    [InlineData(33D, false)]
    [InlineData(0D, true)]
    public void RectangularShapeClipPreservesTransformedPathCoverage(double angle, bool flip) {
        byte[] rectangle = RenderShape(false, angle, flip);
        byte[] path = RenderShape(true, angle, flip);

        Assert.Equal(path, rectangle);
        Assert.Contains(rectangle.Where((_, i) => i % 4 == 3), alpha => alpha > 0);
    }

    private static byte[] RenderGroups(bool usePath, double offset, double scale, int depth) {
        var shape = OfficeShape.Rectangle(14, 14);
        shape.FillColor = OfficeColor.Red;
        var drawing = new OfficeDrawing(24, 24).AddShape(shape, 0, 0);
        for (int i = 0; i < depth; i++) {
            double width = 12.5D - i;
            double height = 11.5D - i;
            drawing = new OfficeDrawing(24, 24).AddClippedDrawing(drawing, offset, offset,
                usePath ? RectangularPath(width, height) : OfficeClipPath.Rectangle(width, height));
        }
        return OfficeDrawingRasterRenderer.Render(drawing, scale).GetPixels();
    }

    private static byte[] RenderShape(bool usePath, double angle, bool flip) {
        var shape = OfficeShape.Rectangle(14, 14);
        shape.FillColor = OfficeColor.SteelBlue;
        shape.ClipPath = usePath ? RectangularPath(10.5D, 9.25D) : OfficeClipPath.Rectangle(10.5D, 9.25D);
        shape.Transform = OfficeTransform.Scale(flip ? -1D : 1D, 1D)
            .Then(OfficeTransform.RotateDegrees(angle))
            .Then(OfficeTransform.Translate(18.5D, 18.25D));
        return OfficeDrawingRasterRenderer.Render(new OfficeDrawing(48, 48).AddShape(shape, 0, 0)).GetPixels();
    }

    private static OfficeClipPath RectangularPath(double width, double height) => OfficeClipPath.Path(
        OfficePathCommand.MoveTo(0D, 0D),
        OfficePathCommand.LineTo(width, 0D),
        OfficePathCommand.LineTo(width, height),
        OfficePathCommand.LineTo(0D, height),
        OfficePathCommand.Close());
}
