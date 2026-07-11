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
            + "<rect width='20' height='20' fill='#00ff00'/><text x='1' y='10'>Pending</text></svg>";

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
    public void SvgReaderConvertsRotatedEllipticalArcsToBoundedCubicPaths() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'>"
            + "<path fill='none' stroke='blue' d='M2 10 A8 6 30 0 1 18 10 A8 6 30 1 1 2 10 Z'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeShape path = Assert.Single(drawing!.Shapes).Shape;
        Assert.Equal(OfficeShapeKind.Path, path.Kind);
        Assert.True(path.PathCommands.Count(command => command.Kind == OfficePathCommandKind.CubicBezierTo) >= 4);
        Assert.Equal(OfficeColor.Blue, path.StrokeColor);
        Assert.Contains("<path", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
        OfficeDrawingRasterRenderer.Render(drawing);
    }

    [Fact]
    public void SvgReaderComposesOrderedNestedTransformsInViewBoxCoordinates() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 20 100 50'>"
            + "<g transform='translate(10 0)'><g transform='scale(2)'><rect x='10' y='20' width='10' height='10' fill='red'/></g></g>"
            + "<rect x='40' y='20' width='10' height='10' fill='blue' transform='translate(10 0) scale(2)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(2, drawing!.Shapes.Count);
        Assert.All(drawing.Shapes, item => Assert.True(item.Shape.Transform.HasValue));
        Assert.Equal(new OfficePoint(20D, 20D), drawing.Shapes[0].Shape.Transform!.Value.TransformPoint(new OfficePoint(0D, 0D)));
        Assert.Equal(new OfficePoint(50D, 20D), drawing.Shapes[1].Shape.Transform!.Value.TransformPoint(new OfficePoint(0D, 0D)));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 25));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(85, 25));
        Assert.Contains("transform=\"matrix(", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderMapsSimpleAnchoredTextToSharedSearchableText() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20' fill='navy' font-family='Arial'>"
            + "<text x='20' y='12' font-size='4' font-weight='700' font-style='italic' text-anchor='middle'>Label</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(drawing!.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Label", text.Text);
        Assert.Equal("Arial", text.Font.FamilyName);
        Assert.True(text.Font.IsBold);
        Assert.True(text.Font.IsItalic);
        Assert.Equal(OfficeColor.Navy, text.Color);
        Assert.True(text.X < 20D);
        Assert.Contains(">Label</text>", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
        OfficeDrawingRasterRenderer.Render(drawing);
    }

    [Fact]
    public void SvgReaderResolvesInheritedLinearPaintServersWithoutRasterFallback() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<defs><linearGradient id='base'><stop offset='20%' stop-color='red'/><stop offset='80%' stop-color='blue'/></linearGradient>"
            + "<linearGradient id='fill' href='#base' x1='0' y1='0.5' x2='1' y2='0.5'/></defs>"
            + "<rect width='40' height='20' fill='url(#fill)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeShape shape = Assert.Single(drawing!.Shapes).Shape;
        OfficeLinearGradient gradient = Assert.IsType<OfficeLinearGradient>(shape.FillGradient);
        Assert.Null(shape.FillColor);
        Assert.Equal(4, gradient.Stops.Count);
        Assert.Equal(0D, gradient.Stops[0].Offset);
        Assert.Equal(1D, gradient.Stops[3].Offset);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        OfficeColor left = raster.GetPixel(2, 10);
        OfficeColor right = raster.GetPixel(37, 10);
        Assert.True(left.R > left.B);
        Assert.True(right.B > right.R);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("<linearGradient", exported, StringComparison.Ordinal);
        Assert.Contains("fill=\"url(#", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderResolvesRadialFillAndLinearStrokePaintServers() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<defs><radialGradient id='spot' cx='50%' cy='50%' r='50%' fx='25%' fy='50%'>"
            + "<stop offset='0' style='stop-color:white;stop-opacity:0.5'/><stop offset='1' stop-color='navy'/></radialGradient>"
            + "<linearGradient id='edge'><stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
            + "<rect x='1' y='1' width='38' height='18' fill='url(#spot)' stroke='url(#edge)' stroke-width='2'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeShape shape = Assert.Single(drawing!.Shapes).Shape;
        OfficeRadialGradient radial = Assert.IsType<OfficeRadialGradient>(shape.FillRadialGradient);
        Assert.IsType<OfficeLinearGradient>(shape.StrokeGradient);
        Assert.InRange(radial.Stops[0].Color.A, (byte)127, (byte)128);
        Assert.Equal(0.25D, radial.StartX);
        Assert.Equal(0.5D, radial.EndRadius);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("<radialGradient", exported, StringComparison.Ordinal);
        Assert.Contains("stroke=\"url(#", exported, StringComparison.Ordinal);
        OfficeDrawingRasterRenderer.Render(drawing);
    }

    [Fact]
    public void SvgReaderDiagnosesUnsafeOrCyclicPaintServersAndKeepsSupportedSiblings() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 10'>"
            + "<defs><linearGradient id='a' href='#b'/><linearGradient id='b' href='#a'/>"
            + "<linearGradient id='duplicate'><stop stop-color='red'/></linearGradient><linearGradient id='duplicate'><stop stop-color='blue'/></linearGradient></defs>"
            + "<rect width='10' height='10' fill='url(#a)'/><rect x='10' width='10' height='10' fill='url(https://example.test/g)'/>"
            + "<rect x='20' width='10' height='10' fill='url(#duplicate)'/><rect x='30' width='10' height='10' fill='lime'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(3, unsupported);
        Assert.Equal(4, drawing!.Shapes.Count);
        Assert.Null(drawing.Shapes[0].Shape.FillColor);
        Assert.Null(drawing.Shapes[1].Shape.FillColor);
        Assert.Null(drawing.Shapes[2].Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, drawing.Shapes[3].Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(35, 5));
    }

    [Fact]
    public void SvgReaderRejectsDocumentsWithDoctypeOrExternalEntities() {
        const string svg = "<!DOCTYPE svg [<!ENTITY xxe SYSTEM 'file:///secret.txt'>]><svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><text>&xxe;</text></svg>";

        Assert.False(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing));
        Assert.Null(drawing);
    }
}
