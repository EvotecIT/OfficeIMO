using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenLineMarkersKeepTheirInteriorUnpaintedOnLinesAndPaths(bool path) {
        OfficeShape shape = path ? OfficeShape.Path(80, 1, new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(80, 0) })
            : OfficeShape.Line(0, 0, 80, 0);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 40, 40);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("<polyline", svg);
        Assert.Contains("fill=\"none\"", svg);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(0, image.GetPixel(75, 45).A);
        Assert.True(image.GetPixel(75, 38).A > 0);
        Assert.True(image.GetPixel(75, 50).A > 0);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 40, 40);
        drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        Assert.Equal(255, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(75, 45).A);
    }

    [Fact]
    public void SvgLineMarkersDoNotPaintWhenTheStrokeIsDisabled() {
        var shape = OfficeShape.Line(0, 0, 80, 0);
        shape.FillColor = OfficeColor.Red;
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 40, 40);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        Assert.DoesNotContain("<polygon", OfficeDrawingSvgExporter.ToSvg(drawing));
        shape.StrokeColor = OfficeColor.Red; shape.StrokeWidth = 0;
        drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        Assert.DoesNotContain("<polygon", OfficeDrawingSvgExporter.ToSvg(drawing));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MarkerGeometryAndOutlineFollowTheCompleteAffineTransform(bool path, bool shear) {
        OfficeShape shape = path ? OfficeShape.Path(70, 40, OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(60, 20))
            : OfficeShape.Line(0, 0, 50, 0);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        OfficeTransform affine = shear ? new OfficeTransform(4, 1, 1, 3, 0, 0) : OfficeTransform.Scale(4, 4);
        shape.Transform = path ? affine : OfficeTransform.Translate(10, 20).Then(affine);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 20, 20);
        var drawing = new OfficeDrawing(300, 180).AddShape(shape, 0, 0);
        OfficeRasterImage open = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(open.GetPixel(shear ? 206 : 180, shear ? 90 : 50).A > 200);
        Assert.True(open.GetPixel(shear ? 206 : 180, shear ? 91 : 53).A > 200);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 20, 20);
        drawing = new OfficeDrawing(300, 180).AddShape(shape, 0, 0);
        OfficeRasterImage filled = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(filled.GetPixel(shear ? 200 : 180, shear ? 100 : 60).A > 200);
    }
}
