using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeDrawingRasterRenderer_GradientUsesPaddedShapeCanvasForPathsAndPolygons(bool radial) {
        OfficeShape path = OfficeShape.Path(200, 100, OfficePathCommand.MoveTo(40, 20), OfficePathCommand.LineTo(160, 20),
            OfficePathCommand.LineTo(160, 80), OfficePathCommand.LineTo(40, 80), OfficePathCommand.Close());
        OfficeShape polygon = OfficeShape.Polygon(new[] { new OfficePoint(0, 0), new OfficePoint(120, 0), new OfficePoint(120, 60), new OfficePoint(0, 60) });
        polygon.Width = 200; polygon.Height = 100;
        OfficeShape reference = OfficeShape.Rectangle(200, 100);
        foreach (OfficeShape shape in new[] { path, polygon, reference }) {
            if (radial) shape.FillRadialGradient = new OfficeRadialGradient(.4, .6, 0, 0, .4, .6, .5, 1,
                new[] { new OfficeGradientStop(0, OfficeColor.Blue), new OfficeGradientStop(1, OfficeColor.Red) });
            else shape.FillGradient = OfficeLinearGradient.DiagonalDown(OfficeColor.Blue, OfficeColor.Red);
            shape.StrokeColor = null;
        }
        var images = new[] { path, reference }.Select(shape => { var drawing = new OfficeDrawing(220, 120); drawing.AddShape(shape, 10, 10); return OfficeDrawingRasterRenderer.Render(drawing); }).ToArray();
        var polygonDrawing = new OfficeDrawing(220, 120); polygonDrawing.AddShape(polygon, 10, 10); var polygonImage = OfficeDrawingRasterRenderer.Render(polygonDrawing);
        foreach (int y in new[] { 35, 60, 85 }) foreach (int x in new[] { 55, 110, 165 }) {
            Assert.Equal(images[1].GetPixel(x, y), images[0].GetPixel(x, y));
        }
        Assert.Equal(images[1].GetPixel(90, 45), polygonImage.GetPixel(90, 45));
        Assert.Equal(0, images[0].GetPixel(25, 25).A);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeDrawingRasterRenderer_AffineRadialFillKeepsItsLocalColorField(bool path) {
        OfficeShape shape = path ? OfficeShape.Path(100, 100, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(100, 0),
            OfficePathCommand.LineTo(100, 100), OfficePathCommand.LineTo(0, 100), OfficePathCommand.Close()) : OfficeShape.Rectangle(100, 100);
        shape.FillRadialGradient = new OfficeRadialGradient(.5, .5, 0, 0, .5, .5, .5, .25,
            new[] { new OfficeGradientStop(0, OfficeColor.Blue), new OfficeGradientStop(1, OfficeColor.Red) }); shape.StrokeColor = null;
        var plain = new OfficeDrawing(300, 150); plain.AddShape(shape, 0, 0);
        shape.Transform = new OfficeTransform(1, 0, 1, 1, 0, 0); var sheared = new OfficeDrawing(300, 150); sheared.AddShape(shape, 0, 0);
        var original = OfficeDrawingRasterRenderer.Render(plain); var transformed = OfficeDrawingRasterRenderer.Render(sheared);
        foreach (var point in new[] { (50, 35), (35, 50), (65, 60) }) {
            var a = original.GetPixel(point.Item1, point.Item2); var b = transformed.GetPixel(point.Item1 + point.Item2, point.Item2);
            Assert.InRange(System.Math.Abs(a.R - b.R), 0, 3); Assert.InRange(System.Math.Abs(a.B - b.B), 0, 3);
        }
    }

    [Fact]
    public void OfficeDrawingRasterRenderer_CompoundEllipticalGradientRetainsHoleAndMatchesRectangle() {
        var drawing = new OfficeDrawing(200, 100); var reference = new OfficeDrawing(200, 100);
        var path = OfficeShape.Path(200, 100, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(200, 0), OfficePathCommand.LineTo(200, 100), OfficePathCommand.LineTo(0, 100), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(60, 30), OfficePathCommand.LineTo(60, 70), OfficePathCommand.LineTo(140, 70), OfficePathCommand.LineTo(140, 30), OfficePathCommand.Close());
        path.FillRadialGradient = new OfficeRadialGradient(.5, .5, 0, 0, .5, .5, .7, 1.4,
            new[] { new OfficeGradientStop(0, OfficeColor.Blue), new OfficeGradientStop(1, OfficeColor.Red) }); path.StrokeColor = null;
        var rect = OfficeShape.Rectangle(200, 100); rect.FillRadialGradient = path.FillRadialGradient; rect.StrokeColor = null;
        drawing.AddShape(path, 0, 0); reference.AddShape(rect, 0, 0); var actual = OfficeDrawingRasterRenderer.Render(drawing); var expected = OfficeDrawingRasterRenderer.Render(reference);
        Assert.Equal(0, actual.GetPixel(100, 50).A);
        foreach (var point in new[] { (100, 15), (30, 50), (170, 50) }) Assert.Equal(expected.GetPixel(point.Item1, point.Item2), actual.GetPixel(point.Item1, point.Item2));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OfficeDrawingRasterRenderer_GradientPathsUseTheCanvasPremultipliedStopInterpolation(bool radial) {
        var path = OfficeShape.Path(100, 100, OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(100, 0), OfficePathCommand.LineTo(100, 100), OfficePathCommand.LineTo(0, 100), OfficePathCommand.Close());
        var rectangle = OfficeShape.Rectangle(100, 100);
        foreach (var shape in new[] { path, rectangle }) {
            shape.StrokeColor = null;
            if (radial) shape.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
            else shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
        }
        var actual = new OfficeDrawing(100, 100); actual.AddShape(path, 0, 0); var expected = new OfficeDrawing(100, 100); expected.AddShape(rectangle, 0, 0);
        var a = OfficeDrawingRasterRenderer.Render(actual); var b = OfficeDrawingRasterRenderer.Render(expected);
        foreach (var point in new[] { (35, 45), (50, 55), (75, 60) }) Assert.Equal(b.GetPixel(point.Item1, point.Item2), a.GetPixel(point.Item1, point.Item2));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void GradientFillTransparentStopsRetainVisibleColorAcrossContours(bool radial, bool transformed) {
        var polygon = OfficeShape.Polygon(new[] { new OfficePoint(0, 0), new OfficePoint(100, 0), new OfficePoint(100, 100), new OfficePoint(0, 100) });
        var compound = OfficeShape.Path(100, 100,
            OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(100, 0), OfficePathCommand.LineTo(100, 100), OfficePathCommand.LineTo(0, 100), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(40, 40), OfficePathCommand.LineTo(40, 60), OfficePathCommand.LineTo(60, 60), OfficePathCommand.LineTo(60, 40), OfficePathCommand.Close());
        foreach (var shape in new[] { polygon, compound }) {
            shape.StrokeColor = null;
            if (radial) shape.FillRadialGradient = OfficeRadialGradient.Centered(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
            else shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
            var transform = transformed ? new OfficeTransform(-1, 0, .5, 1, 120, 0) : OfficeTransform.Identity;
            shape.Transform = transform;
            var drawing = new OfficeDrawing(180, 120); drawing.AddShape(shape, 0, 0);
            var image = OfficeDrawingRasterRenderer.Render(drawing);
            var inverse = transform.Invert();
            foreach (var point in new[] { (25, 35), (75, 65) }) {
                var target = transform.TransformPoint(new OfficePoint(point.Item1, point.Item2));
                int x = (int)target.X, y = (int)target.Y;
                var local = inverse.TransformPoint(new OfficePoint(x + .5, y + .5));
                double ratio = radial ? System.Math.Min(1, System.Math.Sqrt(System.Math.Pow(local.X / 100 - .5, 2) + System.Math.Pow(local.Y / 100 - .5, 2)) / .5) : local.X / 100;
                var color = image.GetPixel(x, y);
                Assert.Equal(255, color.R); Assert.Equal(0, color.G); Assert.Equal(0, color.B);
                Assert.InRange(System.Math.Abs(color.A - 255 * (1 - ratio)), 0, 1);
            }
            if (shape == compound) {
                var hole = transform.TransformPoint(new OfficePoint(50, 50));
                Assert.Equal(0, image.GetPixel((int)hole.X, (int)hole.Y).A);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GradientStrokeRetainsSeparateColorAndAlphaInterpolation(bool radial) {
        var shape = OfficeShape.Path(100, 100, OfficePathCommand.MoveTo(0, 50), OfficePathCommand.LineTo(100, 50));
        shape.FillColor = null; shape.StrokeWidth = 10;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.FromRgba(0, 0, 255, 0)));
        else shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.FromRgba(0, 0, 255, 0));
        var drawing = new OfficeDrawing(100, 100); drawing.AddShape(shape, 0, 0);
        var image = OfficeDrawingRasterRenderer.Render(drawing);
        foreach (int x in new[] { 25, 50, 75 }) {
            double ratio = radial ? System.Math.Sqrt(System.Math.Pow((x + .5) / 100, 2) + System.Math.Pow(.005, 2)) : (x + .5) / 100;
            var color = image.GetPixel(x, 50);
            Assert.InRange(System.Math.Abs(color.R - 255 * (1 - ratio)), 0, 1);
            Assert.InRange(System.Math.Abs(color.B - 255 * ratio), 0, 1);
            Assert.InRange(System.Math.Abs(color.A - 255 * (1 - ratio)), 0, 1);
        }
    }
}
