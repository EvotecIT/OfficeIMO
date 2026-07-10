using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgReaderTests {
    [Fact]
    public void SvgReaderBuildsSharedSceneFromCommonPrimitivesAndInheritedPaint() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 20 100 50'>"
            + "<g fill='red' stroke='blue' stroke-width='2'>"
            + "<rect x='10' y='20' width='20' height='10'/><circle cx='50' cy='30' r='8'/><ellipse cx='75' cy='30' rx='10' ry='5'/>"
            + "<line x1='10' y1='60' x2='110' y2='60'/><polygon points='80,45 100,45 90,60'/></g></svg>";

        bool success = OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported);

        Assert.True(success);
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(100D, drawing!.Width);
        Assert.Equal(50D, drawing.Height);
        Assert.Equal(5, drawing.Shapes.Count);
        Assert.Equal(OfficeColor.Red, drawing.Shapes[0].Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, drawing.Shapes[0].Shape.StrokeColor);
        Assert.Equal(2D, drawing.Shapes[0].Shape.StrokeWidth);
        Assert.Equal((0D, 0D), (drawing.Shapes[0].X, drawing.Shapes[0].Y));
        Assert.Equal(OfficeShapeKind.Line, drawing.Shapes[3].Shape.Kind);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(10, 5));
    }

    [Fact]
    public void SvgReaderRetainsSupportedPrimitivesAndCountsUnsupportedContent() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'>"
            + "<rect width='20' height='20' fill='#00ff00'/><path d='M1 10 A9 9 0 0 1 19 10'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Single(drawing!.Shapes);
        Assert.Equal(1, unsupported);
    }

    [Fact]
    public void SvgReaderParsesAbsoluteRelativeAndSmoothPathCommands() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 20 40 20'>"
            + "<path fill='red' d='M10 20 h20 v20 h-20 z'/>"
            + "<path fill='none' stroke='blue' d='M32 30 C34 20 38 20 40 30 S46 40 48 30 Q46 24 44 30 T40 30'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(2, drawing!.Shapes.Count);
        Assert.All(drawing.Shapes, item => Assert.Equal(OfficeShapeKind.Path, item.Shape.Kind));
        Assert.Contains(drawing.Shapes[1].Shape.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);
        Assert.Contains(drawing.Shapes[1].Shape.PathCommands, command => command.Kind == OfficePathCommandKind.QuadraticBezierTo);
        Assert.Equal(OfficeColor.Red, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(10, 10));
    }

    [Fact]
    public void SvgReaderRejectsDocumentsWithDoctypeOrExternalEntities() {
        const string svg = "<!DOCTYPE svg [<!ENTITY xxe SYSTEM 'file:///secret.txt'>]><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><text>&xxe;</text></svg>";

        Assert.False(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.Null(drawing);
    }
}
