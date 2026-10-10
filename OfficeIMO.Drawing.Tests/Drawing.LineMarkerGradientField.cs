using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OpenFreeformShapeClipUsesItsImplicitlyClosedFill() {
        var commands = new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(40, 0), OfficePathCommand.LineTo(40, 40) };
        var shape = OfficeShape.Rectangle(40, 40);
        shape.FillColor = OfficeColor.Red; shape.ClipPath = OfficeClipPath.Path(commands);
        OfficeRasterImage open = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(60, 60).AddShape(shape, 10, 10));
        shape.ClipPath = OfficeClipPath.Path(commands.Concat(new[] { OfficePathCommand.Close() }));
        OfficeRasterImage closed = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(60, 60).AddShape(shape, 10, 10));
        Assert.Equal(closed.GetPixel(45, 20), open.GetPixel(45, 20));
        Assert.Equal(255, open.GetPixel(45, 20).A);
        Assert.Equal(0, open.GetPixel(15, 45).A);
        Assert.True(OfficeTextInkClip.TryCreatePath(OfficeClipPath.Path(commands), OfficeTransform.Identity, default, out _));
    }

    [Fact]
    public void SvgMarkerGradientKeepsNarrowTransitionCoordinates() {
        OfficeShape shape = GradientMarkerShape(false, false);
        shape.StrokeGradient = new OfficeLinearGradient(.4996, .5, .5004, .5,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        var xml = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(240, 140).AddShape(shape, 20, 20)));
        XNamespace ns = "http://www.w3.org/2000/svg";
        var gradient = Assert.Single(xml.Descendants(ns + "linearGradient"));
        Assert.Equal("0.4996", (string?)gradient.Attribute("x1"));
        Assert.Equal("0.5004", (string?)gradient.Attribute("x2"));
    }

    [Fact]
    public void CollapsedGradientLineKeepsProjectedOpenMarker() {
        var shape = OfficeShape.Line(0, 0, 80, 50);
        shape.StrokeWidth = 4;
        shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 100, 10);
        shape.Transform = OfficeTransform.Scale(1, 0);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(150, 80).AddShape(shape, 20, 20));
        Assert.True(image.GetPixel(110, 20).A > 0);
        Assert.True(image.GetPixel(110, 20).B > image.GetPixel(110, 20).R);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SvgMarkersShareTheShaftGradientCoordinateField(bool radial, bool transformed) {
        OfficeShape shape = GradientMarkerShape(radial, transformed);
        var xml = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(240, 140).AddShape(shape, 20, 20)));
        XNamespace ns = "http://www.w3.org/2000/svg";
        var gradient = Assert.Single(xml.Descendants(ns + (radial ? "radialGradient" : "linearGradient")));
        Assert.Equal("userSpaceOnUse", (string?)gradient.Attribute("gradientUnits"));
        Assert.Equal(transformed ? "matrix(80 0 0 40 0 0)" : "matrix(80 0 0 40 20 20)",
            (string?)gradient.Attribute("gradientTransform"));
        string paint = "url(#" + (string?)gradient.Attribute("id") + ")";
        Assert.Equal(paint, (string?)Assert.Single(xml.Descendants(ns + "path")).Attribute("stroke"));
        Assert.Equal(paint, (string?)Assert.Single(xml.Descendants(ns + "polygon")).Attribute("fill"));
        Assert.Equal(paint, (string?)Assert.Single(xml.Descendants(ns + "polyline")).Attribute("stroke"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RasterMarkersSampleTheGradientAcrossTheirPaintedArea(bool radial, bool transformed) {
        OfficeShape shape = GradientMarkerShape(radial, transformed);
        var reference = OfficeShape.Rectangle(80, 40);
        reference.FillGradient = shape.StrokeGradient;
        reference.FillRadialGradient = shape.StrokeRadialGradient;
        reference.Transform = shape.Transform;
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(240, 140).AddShape(shape, 20, 20));
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(240, 140).AddShape(reference, 20, 20));
        int scale = transformed ? 2 : 1;
        OfficeColor painted = actual.GetPixel(20 + 25 * scale, 20 + 20 * scale);
        OfficeColor field = expected.GetPixel(20 + 25 * scale, 20 + 20 * scale);
        Assert.Equal(255, painted.A);
        Assert.InRange((int)painted.R, field.R - 2, field.R + 2);
        Assert.InRange((int)painted.B, field.B - 2, field.B + 2);
        Assert.True(painted.B > 20);
    }

    private static OfficeShape GradientMarkerShape(bool radial, bool transformed) {
        var shape = OfficeShape.Path(80, 40, OfficePathCommand.MoveTo(0, 20), OfficePathCommand.LineTo(80, 20));
        shape.StrokeWidth = 2;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 40, 40);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 40, 40);
        if (transformed) shape.Transform = OfficeTransform.Scale(2, 2);
        return shape;
    }
}
