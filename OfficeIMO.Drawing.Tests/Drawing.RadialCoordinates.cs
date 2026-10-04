using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRadialCoordinatesTests {
    [Theory]
    [InlineData(false, .25, 127)]
    [InlineData(true, .25, 127)]
    [InlineData(false, 0, 191)]
    [InlineData(true, 0, 191)]
    public void ShrinkingCirclesUseTheNonnegativeRadiusRoot(bool stroke, double endRadius, int expectedRed) {
        var gradient = new OfficeRadialGradient(.5, .5, .5, .5, .5, endRadius,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var shape = stroke
            ? OfficeShape.Path(100, 100, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(0, 100),
                OfficePathCommand.MoveTo(0, 50), OfficePathCommand.LineTo(100, 50))
            : OfficeShape.Rectangle(100, 100);
        shape.FillColor = null;
        shape.StrokeWidth = stroke ? 10 : 0;
        if (stroke) shape.StrokeRadialGradient = gradient;
        else shape.FillRadialGradient = gradient;
        var image = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(100, 100).AddShape(shape, 0, 0));
        var color = image.GetPixel(87, 50);
        Assert.InRange((int)color.R, expectedRed - 3, expectedRed + 3);
        Assert.InRange((int)color.B, 255 - expectedRed - 3, 255 - expectedRed + 3);
    }

    [Fact]
    public void CloningAndFillOpacityPreserveTheAffineEllipseField() {
        var transform = OfficeTransform.RotateDegrees(31, .5, .5).Then(new OfficeTransform(1, .2, .3, 1, -.1, -.1));
        var original = new OfficeRadialGradient(.5, .5, 0, 0, .5, .5, .35, .15,
            new[] { new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue) });
        var gradient = original.TransformCoordinates(transform).Clone();
        var shape = OfficeShape.Rectangle(100, 100); shape.FillRadialGradient = gradient; shape.FillOpacity = .4; shape.StrokeWidth = 0;
        var image = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(100, 100).AddShape(shape, 0, 0), background: OfficeColor.White);
        var inverse = transform.Invert();
        foreach (var xy in new[] { (45, 45), (50, 60), (65, 55) }) {
            var point = inverse.TransformPoint(new OfficePoint((xy.Item1 + .5) / 100, (xy.Item2 + .5) / 100));
            double ratio = Math.Min(1, Math.Sqrt(Math.Pow((point.X - .5) / .35, 2) + Math.Pow((point.Y - .5) / .15, 2)));
            Assert.InRange(Math.Abs(image.GetPixel(xy.Item1, xy.Item2).R - (153 + 102 * (1 - ratio))), 0, 2);
        }
        Assert.Equal(OfficeTransform.Identity, original.CoordinateTransform);
        Assert.Throws<ArgumentException>(() => original.TransformCoordinates(OfficeTransform.Scale(0, 1)));
    }
}
