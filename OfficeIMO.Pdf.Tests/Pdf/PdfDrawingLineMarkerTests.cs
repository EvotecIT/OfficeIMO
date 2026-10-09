using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfDrawingLineMarkerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DrawingMarkersReachPdfAsFilledAndOpenGeometryWithStrokeOpacity(bool path) {
        OfficeColor ink = OfficeColor.FromRgb(204, 32, 64);
        OfficeShape shape = path ? OfficeShape.Path(80, 1, new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(40, 0), OfficePathCommand.LineTo(80, 0) })
            : OfficeShape.Line(0, 0, 80, 0);
        shape.StrokeColor = ink; shape.StrokeWidth = 2; shape.StrokeOpacity = 0.5;
        if (path) shape.FillColor = OfficeColor.Blue;
        shape.FillOpacity = 0.1;
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 10, 14);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 18, 20);
        shape.Transform = OfficeTransform.RotateDegrees(30, 40, 0);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 50);
        PdfDocument pdf = PdfDocument.Create().Compose(builder => builder.Page(page => page.Size(140, 100).Margin(0)
            .Canvas(canvas => canvas.Drawing(drawing, 0, 0, 140, 100))));
        OfficeDrawing read = PdfReadDocument.Open(pdf.ToBytes()).Pages[0].ToDrawing();
        OfficeShape[] painted = Shapes(read).Select(item => item.Shape).ToArray();
        OfficeShape marker = Assert.Single(painted, item => item.FillColor == ink);
        Assert.Equal(0.5, marker.FillOpacity.GetValueOrDefault(1), 5);
        Assert.Equal(2, painted.Count(item => item.StrokeColor == ink));
        Assert.All(painted.Where(item => item.StrokeColor == ink), item => Assert.Equal(0.5, item.StrokeOpacity.GetValueOrDefault(1), 5));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GradientMarkersRetainStrokeBrushOpacityAndClipWithoutRasterFallback(bool radial) {
        var shape = OfficeShape.Path(80, 40, OfficePathCommand.MoveTo(10, 20), OfficePathCommand.LineTo(70, 20));
        shape.StrokeWidth = 4;
        shape.StrokeOpacity = .5;
        if (radial) shape.StrokeRadialGradient = new OfficeRadialGradient(0, .5, 0, 0, .5, 1,
            new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue));
        else shape.StrokeGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.StrokeStartMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 20, 20);
        shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Arrow, 20, 20);
        shape.ClipPath = OfficeClipPath.Rectangle(60, 40);
        var drawing = new OfficeDrawing(140, 100).AddShape(shape, 20, 20);
        byte[] bytes = PdfDocument.Create().Compose(builder => builder.Page(page => page.Size(140, 100).Margin(0)
            .Canvas(canvas => canvas.Drawing(drawing, 0, 0, 140, 100)))).ToBytes();
        Assert.DoesNotContain("/Subtype /Image", System.Text.Encoding.ASCII.GetString(bytes));
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes));
        OfficeColor filledEnd = image.GetPixel(43, 36), openEnd = image.GetPixel(75, 33);
        Assert.InRange(filledEnd.A, (byte)120, (byte)135);
        Assert.True(filledEnd.R > filledEnd.B);
        Assert.InRange(openEnd.A, (byte)120, (byte)135);
        Assert.True(openEnd.B > openEnd.R);
        Assert.Equal(0, image.GetPixel(76, 36).A);
        Assert.Equal(0, image.GetPixel(84, 37).A);
        Assert.Null(shape.FillGradient);
        Assert.Null(shape.FillRadialGradient);
        Assert.Equal(.5, shape.StrokeOpacity);
    }

    private static IEnumerable<OfficeDrawingShape> Shapes(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingShape shape) yield return shape;
            if (element is OfficeDrawingGroup group) foreach (OfficeDrawingShape child in Shapes(group.InnerDrawing)) yield return child;
            if (element is OfficeDrawingEffectGroup effect) foreach (OfficeDrawingShape child in Shapes(effect.InnerDrawing)) yield return child;
        }
    }
}
