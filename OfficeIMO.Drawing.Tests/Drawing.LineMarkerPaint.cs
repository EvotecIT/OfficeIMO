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
}
