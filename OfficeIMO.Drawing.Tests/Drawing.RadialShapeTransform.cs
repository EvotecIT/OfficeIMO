using System;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingRadialShapeTransformTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RadialFillFollowsShapeShearAndReflection(bool path, bool reflection) {
        var drawing = new OfficeDrawing(200, 160);
        var shape = path ? OfficeShape.Path(80, 60, new[] {
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(80, 0),
            OfficePathCommand.LineTo(80, 60), OfficePathCommand.LineTo(0, 60), OfficePathCommand.Close()
        }) : OfficeShape.Rectangle(80, 60);
        shape.FillRadialGradient = new OfficeRadialGradient(.3, .4, 0, .3, .4, .5,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var transform = reflection ? new OfficeTransform(-1, .2, .3, 1, 100, 0) : new OfficeTransform(1, .2, .3, 1, 0, 0);
        shape.Transform = transform; drawing.AddShape(shape, 20, 20);
        var raster = OfficeDrawingRasterRenderer.Render(drawing);
        var inverse = transform.Invert();
        for (int y = 35; y < 75; y += 10) for (int x = 50; x < 95; x += 10) {
            var p = inverse.TransformPoint(new OfficePoint(x + .5 - 20, y + .5 - 20));
            if (p.X < 3 || p.X > 77 || p.Y < 3 || p.Y > 57) continue;
            double ratio = Math.Min(1, Math.Sqrt(Math.Pow(p.X / 80 - .3, 2) + Math.Pow(p.Y / 60 - .4, 2)) / .5);
            Assert.InRange(Math.Abs(raster.GetPixel(x, y).R - 255 * (1 - ratio)), 0, 2);
            Assert.InRange(Math.Abs(raster.GetPixel(x, y).B - 255 * ratio), 0, 2);
        }
    }
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GradientStrokeFollowsShapeReflection(bool reflection, bool linear) {
        var drawing = new OfficeDrawing(200, 160);
        var shape = OfficeShape.Path(80, 50, new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(80, 50) });
        shape.FillColor = null; shape.StrokeWidth = 10;
        shape.StrokeRadialGradient = new OfficeRadialGradient(.3, .4, 0, .3, .4, .5,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var transform = reflection ? new OfficeTransform(-1, .2, .3, 1, 100, 0) : new OfficeTransform(1, .2, .3, 1, 0, 0);
        if (linear) { shape.StrokeRadialGradient = null; shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue); }
        shape.Transform = transform; drawing.AddShape(shape, 20, 20);
        var image = OfficeDrawingRasterRenderer.Render(drawing); var inverse = transform.Invert();
        foreach (double t in new[] { .2, .4, .6, .8 }) {
            var destination = transform.TransformPoint(new OfficePoint(80 * t, 50 * t));
            int x = (int)(destination.X + 20), y = (int)(destination.Y + 20);
            var local = inverse.TransformPoint(new OfficePoint(x + .5 - 20, y + .5 - 20));
            double ratio = linear ? local.X / 80 : Math.Min(1, Math.Sqrt(Math.Pow(local.X / 80 - .3, 2) + Math.Pow(local.Y / 50 - .4, 2)) / .5);
            Assert.InRange(Math.Abs(image.GetPixel(x, y).R - 255 * (1 - ratio)), 0, 2);
        }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReflectedGradientMarkersKeepEndpointColors(bool path) {
        var drawing = new OfficeDrawing(150, 100);
        var shape = path ? OfficeShape.Path(80, 50, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(80, 50))
            : OfficeShape.Line(0, 0, 80, 50);
        shape.StrokeWidth = 2; shape.FillColor = null;
        shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Diamond, 12, 12);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Diamond, 12, 12);
        shape.Transform = new OfficeTransform(-1, 0, 0, 1, 100, 0);
        drawing.AddShape(shape, 10, 20);
        var image = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, image.GetPixel(109, 20));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(30, 69));
    }
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public void InsetDeclaredCanvasKeepsPaintWhenTranslated(bool stroke, bool linear, bool markers) {
        var commands = stroke ? new[] { OfficePathCommand.MoveTo(20, 20), OfficePathCommand.LineTo(80, 80) }
            : new[] { OfficePathCommand.MoveTo(20, 20), OfficePathCommand.LineTo(80, 20),
                OfficePathCommand.LineTo(80, 80), OfficePathCommand.LineTo(20, 80), OfficePathCommand.Close() };
        var shape = OfficeShape.Path(100, 100, commands);
        var radial = new OfficeRadialGradient(.3, .4, 0, .3, .4, .5,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var horizontal = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.FillColor = null;
        if (stroke) {
            shape.StrokeWidth = 8;
            if (linear) shape.StrokeGradient = horizontal; else shape.StrokeRadialGradient = radial;
            if (markers) {
                shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Diamond, 12, 12);
                shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Diamond, 12, 12);
            }
        } else if (linear) shape.FillGradient = horizontal; else shape.FillRadialGradient = radial;
        var before = new OfficeDrawing(120, 120); before.AddShape(shape, 10, 10);
        shape.Transform = OfficeTransform.Translate(1, 0);
        var after = new OfficeDrawing(120, 120); after.AddShape(shape, 10, 10);
        var first = OfficeDrawingRasterRenderer.Render(before, background: OfficeColor.White);
        var second = OfficeDrawingRasterRenderer.Render(after, background: OfficeColor.White);
        for (int y = 35; y < 85; y++) for (int x = 35; x < 85; x++) {
            if (stroke && Math.Abs(x - y) > 1) continue;
            Assert.InRange(Math.Abs(first.GetPixel(x, y).R - second.GetPixel(x + 1, y).R), 0, 2);
            Assert.InRange(Math.Abs(first.GetPixel(x, y).B - second.GetPixel(x + 1, y).B), 0, 2);
        }
        if (markers) foreach (int coordinate in new[] { 30, 90 }) {
            Assert.Equal(first.GetPixel(coordinate, coordinate), second.GetPixel(coordinate + 1, coordinate));
        }
        double ratio = linear ? .305 : Math.Sqrt(Math.Pow(.305 - .3, 2) + Math.Pow(.305 - .4, 2)) / .5;
        Assert.InRange(Math.Abs(first.GetPixel(40, 40).R - 255 * (1 - ratio)), 0, 2);
    }
}
