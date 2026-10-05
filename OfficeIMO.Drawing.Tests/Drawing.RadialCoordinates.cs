using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRadialCoordinatesTests {
    [Theory]
    [InlineData(1.000000005)]
    [InlineData(1D)]
    [InlineData(.999999995)]
    public void NearlyLinearShrinkingFieldsRetainThePhysicalRoot(double focus) {
        var field = new OfficeRadialGradient(0, 0, 1, focus, 0, 0,
            new OfficeGradientStop(0, OfficeColor.Blue), new OfficeGradientStop(1, OfficeColor.Red));
        // A small painted region beside the focus remains inside a physical
        // circle even when both quadratic coefficients are below 1e-7.
        Assert.InRange(field.SampleRatio(focus - .0000000025, .000000025), .99999, 1D);
    }

    [Fact]
    public void NearPointEndDoesNotRoundOutsideRootsOntoAZeroRadiusCircle() {
        var field = new OfficeRadialGradient(0, 0, 1, 1.000000005, 0, 0,
            new OfficeGradientStop(0, OfficeColor.Blue), new OfficeGradientStop(1, OfficeColor.Red));
        Assert.Equal(0D, field.SampleRatio(1.0000000055, -.0000000105));
    }

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

    [Fact]
    public void NativePadEndpointPaintTracksStopAlphaThroughCloneAndOpacity() {
        var field = new OfficeRadialGradient(1, .5, 0, 0, .5, .5, .3, .2,
            new[] { new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue) })
            .WithFirstPadIntersection().Clone().WithStops(new[] {
                new OfficeGradientStop(0, OfficeColor.FromRgba(0, 0, 255, 64)),
                new OfficeGradientStop(1, OfficeColor.FromRgba(255, 0, 0, 128)) });
        Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 64), field.OutsideColor);
        var shape = OfficeShape.Rectangle(100, 100);
        shape.FillRadialGradient = field; shape.FillOpacity = .5; shape.StrokeWidth = 0;
        var image = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(100, 100).AddShape(shape, 0, 0), background: OfficeColor.White);
        Assert.InRange(image.GetPixel(95, 5).R, 221, 225); Assert.Equal(255, image.GetPixel(95, 5).B);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgRejectsNativeEndpointFieldsInFillAndStroke(bool stroke) {
        var field = new OfficeRadialGradient(1, .5, 0, .5, .5, .3,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue)).WithFirstPadIntersection();
        var shape = OfficeShape.Rectangle(100, 100);
        if (stroke) shape.StrokeRadialGradient = field;
        else shape.FillRadialGradient = field;
        var drawing = new OfficeDrawing(100, 100).AddShape(shape, 0, 0);
        Assert.Throws<NotSupportedException>(() => OfficeDrawingSvgExporter.ToSvg(drawing));
    }
}
