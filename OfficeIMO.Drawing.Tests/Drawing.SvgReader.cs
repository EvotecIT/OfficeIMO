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
    public void SvgReaderResolvesPercentageGeometryAgainstTheViewport() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 20 200 100'>"
            + "<rect x='10%' y='20%' width='50%' height='25%' fill='red'/>"
            + "<circle cx='75%' cy='50%' r='10%' fill='blue'/>"
            + "<ellipse cx='25%' cy='75%' rx='5%' ry='10%' fill='green'/>"
            + "<line x1='10%' y1='90%' x2='90%' y2='30%' stroke='black'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(4, drawing!.Shapes.Count);

        OfficeDrawingShape rectangle = drawing.Shapes[0];
        Assert.Equal(10D, rectangle.X, 6);
        Assert.Equal(0D, rectangle.Y, 6);
        Assert.Equal(100D, rectangle.Shape.Width, 6);
        Assert.Equal(25D, rectangle.Shape.Height, 6);

        double circleRadius = Math.Sqrt((200D * 200D) + (100D * 100D)) / Math.Sqrt(2D) * 0.1D;
        OfficeDrawingShape circle = drawing.Shapes[1];
        Assert.Equal(140D - circleRadius, circle.X, 6);
        Assert.Equal(30D - circleRadius, circle.Y, 6);
        Assert.Equal(circleRadius * 2D, circle.Shape.Width, 6);

        OfficeDrawingShape ellipse = drawing.Shapes[2];
        Assert.Equal(30D, ellipse.X, 6);
        Assert.Equal(45D, ellipse.Y, 6);
        Assert.Equal(20D, ellipse.Shape.Width, 6);
        Assert.Equal(20D, ellipse.Shape.Height, 6);
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
    public void SvgReaderAllowsAnExplicitTrustedElementLimit() {
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 1 1'>");
        for (int i = 0; i < OfficeSvgDrawingReaderOptions.DefaultMaximumElements + 1; i++) svg.Append("<metadata/>");
        svg.Append("</svg>");
        byte[] bytes = Encoding.UTF8.GetBytes(svg.ToString());

        Assert.False(OfficeSvgDrawingReader.TryRead(bytes, out _));

        var options = new OfficeSvgDrawingReaderOptions {
            MaximumElements = OfficeSvgDrawingReaderOptions.DefaultMaximumElements + 1
        };
        Assert.True(OfficeSvgDrawingReader.TryRead(bytes, options, out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Empty(drawing!.Elements);
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
    public void SvgReaderPreservesThePathBudgetAfterMalformedPathData() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 20'>"
            + "<path d='M0 0 L nope' stroke='red'/>"
            + "<path d='M1 1 L10 10' stroke='blue'/>"
            + "<polyline points='1,10 5,15 10,10' stroke='green'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        Assert.Equal(2, drawing!.Shapes.Count);
        Assert.Contains(drawing.Shapes, item => item.Shape.StrokeColor == OfficeColor.Blue);
        Assert.Contains(drawing.Shapes, item => item.Shape.StrokeColor == OfficeColor.Green);
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
    public void SvgReaderResolvesPercentageTextPositionAndHangingBaseline() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='10 20 200 100' fill='navy'>"
            + "<text x='50%' y='25%' font-size='10' text-anchor='middle' dominant-baseline='hanging'>Label</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingText text = Assert.Single(drawing!.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Label", text.Text);
        Assert.Equal(90D, text.X + (text.Width / 2D), 6);
        Assert.Equal(5D, text.Y, 6);
        Assert.Contains(">Label</text>", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderPreservesRgbAndRgbaPaint() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'>"
            + "<rect width='10' height='10' fill='rgba(36,87,166,0.502)'/>"
            + "<rect x='10' width='10' height='10' fill='rgb(10%,20%,30%)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(OfficeColor.FromRgba(36, 87, 166, 128), drawing!.Shapes[0].Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgb(26, 51, 76), drawing.Shapes[1].Shape.FillColor);
    }

    [Fact]
    public void SvgReaderKeepsAffineTransformedTextAsNativeDrawingText() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20' fill='navy'>"
            + "<text x='4' y='12' font-size='5' transform='translate(3 -1) skewX(8)'>AffineLabel</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup group = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingText text = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("AffineLabel", text.Text);
        Assert.NotEqual(OfficeTransform.Identity, group.Transform);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("transform=\"matrix(", exported, StringComparison.Ordinal);
        Assert.Contains(">AffineLabel</text>", exported, StringComparison.Ordinal);
        OfficeDrawingRasterRenderer.Render(drawing);
    }

    [Fact]
    public void SvgReaderMapsNestedTspanRunsWithInheritedStyleAndPositioning() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 50 24' fill='navy' font-family='Arial'>"
            + "<text x='25' y='10' font-size='4' text-anchor='middle'>One<tspan fill='red' font-weight='bold'> Two</tspan>"
            + "<tspan x='5' y='20' dx='2' font-style='italic' text-anchor='start'>Three</tspan></text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingText[] runs = drawing!.Elements.OfType<OfficeDrawingText>().ToArray();
        Assert.Equal(new[] { "One", " Two", "Three" }, runs.Select(run => run.Text));
        Assert.Equal(OfficeColor.Navy, runs[0].Color);
        Assert.Equal(OfficeColor.Red, runs[1].Color);
        Assert.True(runs[1].Font.IsBold);
        Assert.True(runs[2].Font.IsItalic);
        Assert.True(runs[0].X < 25D);
        Assert.Equal(7D, runs[2].X, 3);
        Assert.Contains(">One</text>", OfficeDrawingSvgExporter.ToSvg(drawing), StringComparison.Ordinal);
        OfficeDrawingRasterRenderer.Render(drawing);
    }

    [Fact]
    public void SvgReaderScalesTextLengthThroughSearchableEffectGroups() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<text x='2' y='14' font-size='8' textLength='24' lengthAdjust='spacingAndGlyphs'>Wide</text></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeDrawingEffectGroup group = Assert.Single(drawing!.Elements.OfType<OfficeDrawingEffectGroup>());
        OfficeDrawingText text = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal("Wide", text.Text);
        Assert.True(group.Transform.M11 > 1D);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("transform=\"matrix(", exported, StringComparison.Ordinal);
        Assert.Contains(">Wide</text>", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderResolvesInheritedCurrentColorForShapePaint() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 30 10'>"
            + "<g color='purple'><rect width='10' height='10' fill='currentColor'/>"
            + "<line x1='10' y1='5' x2='20' y2='5' stroke='currentColor' stroke-width='2'/></g>"
            + "<rect x='20' width='10' height='10' style='fill:currentColor;color:lime'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(3, drawing!.Shapes.Count);
        Assert.Equal(OfficeColor.Purple, drawing.Shapes[0].Shape.FillColor);
        Assert.Equal(OfficeColor.Purple, drawing.Shapes[1].Shape.StrokeColor);
        Assert.Equal(OfficeColor.Lime, drawing.Shapes[2].Shape.FillColor);
    }

    [Fact]
    public void SvgReaderExpandsBoundedLocalUseReferencesWithInheritedPaintAndPlacement() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink' viewBox='0 0 40 20'>"
            + "<defs><g id='badge'><rect width='10' height='10'/><circle cx='5' cy='5' r='3' fill='white'/></g></defs>"
            + "<use href='#badge' x='2' y='2' fill='red'/><use xlink:href='#badge' x='22' y='5' fill='blue'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(4, drawing!.Shapes.Count);
        Assert.Equal(OfficeColor.Red, drawing.Shapes[0].Shape.FillColor);
        Assert.Equal(OfficeColor.Blue, drawing.Shapes[2].Shape.FillColor);
        Assert.All(drawing.Shapes, item => Assert.True(item.Shape.Transform.HasValue));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(3, 3));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(23, 6));
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.DoesNotContain("<use", exported, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderMapsLocalSymbolViewportsThroughSharedEffectGroups() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<defs><symbol id='badge' viewBox='0 0 10 10'><rect width='10' height='10'/></symbol></defs>"
            + "<use href='#badge' x='2' y='2' width='20' height='10' fill='red'/>"
            + "<use href='#badge' x='24' y='4' width='12' height='8' preserveAspectRatio='none' fill='blue'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(2, drawing!.Elements.OfType<OfficeDrawingEffectGroup>().Count());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(8, 5));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(26, 6));
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.DoesNotContain("<symbol", exported, StringComparison.Ordinal);
        Assert.DoesNotContain("<use", exported, StringComparison.Ordinal);
        Assert.True(exported.Split(new[] { "transform=\"matrix(" }, StringSplitOptions.None).Length >= 3);
    }

    [Fact]
    public void SvgReaderAlignsAndClipsMeetAndSliceSymbolViewports() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 12'>"
            + "<defs><symbol id='badge' viewBox='0 0 10 10'><rect x='0' width='5' height='10' fill='red'/><rect x='5' width='5' height='10' fill='blue'/></symbol></defs>"
            + "<use href='#badge' width='12' height='8' preserveAspectRatio='xMaxYMid meet'/>"
            + "<use href='#badge' x='16' width='12' height='8' preserveAspectRatio='xMinYMid slice'/>"
            + "<use href='#badge' x='32' width='8' height='8' preserveAspectRatio='defer xMidYMid meet'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(1, 4));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 4));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(16, 4));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(27, 4));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(39, 4));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(28, 4));
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("clip-path=\"url(#officeimo-group-clip-", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderDiagnosesCyclicExternalAndAmbiguousUseReferences() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 10'>"
            + "<defs><g id='loop'><use href='#loop'/></g><rect id='duplicate' width='5' height='5'/><circle id='duplicate' r='2'/></defs>"
            + "<use href='#loop'/><use href='https://example.test/shape'/><use href='#duplicate'/><rect x='30' width='10' height='10' fill='lime'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(3, unsupported);
        OfficeDrawingShape shape = Assert.Single(drawing!.Shapes);
        Assert.Equal(OfficeColor.Lime, shape.Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, OfficeDrawingRasterRenderer.Render(drawing).GetPixel(35, 5));
    }

    [Fact]
    public void SvgReaderBoundsNestedTextTransformsAndSymbolSurfaces() {
        var nested = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><text>");
        for (int index = 0; index < 160; index++) nested.Append("<tspan>");
        nested.Append("Text");
        for (int index = 0; index < 160; index++) nested.Append("</tspan>");
        nested.Append("</text><line x2='10' y2='10' stroke='black' stroke-dasharray='1 1' transform='scale(1000000000)'/></svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(nested.ToString()),
            out OfficeDrawing? bounded, out int unsupported));
        Assert.NotNull(bounded);
        Assert.True(unsupported >= 2);
        OfficeDrawingRasterRenderer.Render(bounded!);

        const string oversized = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>"
            + "<defs><symbol id='large' viewBox='0 0 100000 100000'><rect width='1' height='1'/></symbol></defs>"
            + "<use href='#large' width='100000' height='100000'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(oversized),
            out OfficeDrawing? safe, out int surfaceUnsupported));
        Assert.NotNull(safe);
        Assert.Empty(safe!.Elements);
        Assert.Equal(1, surfaceUnsupported);
    }

    [Fact]
    public void SvgReaderChargesRepeatedUsePathsToOneCommandBudget() {
        var path = new StringBuilder("M0 0");
        for (int index = 0; index < 1000; index++) path.Append(" L1 1");
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs><path id='p' d='")
            .Append(path).Append("' stroke='black'/></defs>");
        for (int index = 0; index < 25; index++) svg.Append("<use href='#p'/>");
        svg.Append("</svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.ToString()),
            out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.InRange(drawing!.Shapes.Count, 1, 19);
        Assert.True(unsupported > 0);
        Assert.True(drawing.Shapes.Sum(shape => shape.Shape.PathCommands.Count) <= 20000);
    }

    [Fact]
    public void SvgReaderChargesRepeatedUsePolylinesToOneCommandBudget() {
        var points = new StringBuilder("0,0");
        for (int index = 1; index < 1000; index++) points.Append(' ').Append(index % 10).Append(',').Append(index % 10);
        var svg = new StringBuilder("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'><defs><polyline id='p' points='")
            .Append(points).Append("' stroke='black'/></defs>");
        for (int index = 0; index < 25; index++) svg.Append("<use href='#p'/>");
        svg.Append("</svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg.ToString()),
            out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.InRange(drawing!.Shapes.Count, 1, 20);
        Assert.True(unsupported > 0);
        Assert.True(drawing.Shapes.Sum(shape => shape.Shape.PathCommands.Count) <= 20000);
    }

    [Fact]
    public void SvgReaderDoesNotExhaustPathBudgetForMalformedPointLists() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'>"
            + "<polygon points='0,0 nope'/><polyline points='0,0 1'/><path d='M10 0 L20 10' stroke='lime'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(2, unsupported);
        OfficeDrawingShape shape = Assert.Single(drawing!.Shapes);
        Assert.Equal(OfficeColor.Lime, shape.Shape.StrokeColor);
        Assert.Equal(2, shape.Shape.PathCommands.Count);
    }

    [Fact]
    public void SvgReaderRejectsTransformArgumentsBeyondSupportedArity() {
        var arguments = new StringBuilder("0");
        for (int index = 1; index < 1000; index++) arguments.Append(" 0");
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10'>"
            + "<rect width='10' height='10' fill='red' transform='matrix(" + arguments + ")'/>"
            + "<rect x='10' width='10' height='10' fill='lime'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        Assert.Equal(2, drawing!.Shapes.Count);
        Assert.Equal(OfficeColor.Red, drawing.Shapes[0].Shape.FillColor);
        Assert.False(drawing.Shapes[0].Shape.Transform.HasValue);
        Assert.Equal(OfficeColor.Lime, drawing.Shapes[1].Shape.FillColor);
    }

    [Fact]
    public void SvgReaderResolvesInheritedLinearPaintServersWithoutRasterFallback() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'>"
            + "<defs><linearGradient id='base'><stop offset='20%' stop-color='red'/><stop offset='50%' stop-color='red'/><stop offset='50%' stop-color='blue'/><stop offset='80%' stop-color='blue'/></linearGradient>"
            + "<linearGradient id='fill' href='#base' x1='0' y1='0.5' x2='1' y2='0.5'/></defs>"
            + "<rect width='40' height='20' fill='url(#fill)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeShape shape = Assert.Single(drawing!.Shapes).Shape;
        OfficeLinearGradient gradient = Assert.IsType<OfficeLinearGradient>(shape.FillGradient);
        Assert.Null(shape.FillColor);
        Assert.Equal(6, gradient.Stops.Count);
        Assert.Equal(0D, gradient.Stops[0].Offset);
        Assert.Equal(0.5D, gradient.Stops[2].Offset);
        Assert.Equal(0.5D, gradient.Stops[3].Offset);
        Assert.Equal(1D, gradient.Stops[5].Offset);

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
    public void SvgReaderResolvesUserSpacePaintServersPerTargetShape() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 40 20'><defs>"
            + "<linearGradient id='shared' gradientUnits='userSpaceOnUse' x1='0' y1='0' x2='50%' y2='0'>"
            + "<stop offset='0' stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<radialGradient id='spot' gradientUnits='userSpaceOnUse' cx='30' cy='10' r='8' fx='28' fy='10'>"
            + "<stop offset='0' stop-color='white'/><stop offset='1' stop-color='navy'/></radialGradient>"
            + "</defs><rect width='10' height='20' fill='url(#shared)'/><rect x='10' width='10' height='20' fill='url(#shared)'/>"
            + "<rect x='20' width='20' height='20' fill='url(#spot)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.Equal(3, drawing!.Shapes.Count);
        OfficeLinearGradient first = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[0].Shape.FillGradient);
        OfficeLinearGradient second = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[1].Shape.FillGradient);
        Assert.Equal(0D, first.StartX);
        Assert.Equal(2D, first.EndX);
        Assert.Equal(-1D, second.StartX);
        Assert.Equal(1D, second.EndX);
        OfficeRadialGradient radial = Assert.IsType<OfficeRadialGradient>(drawing.Shapes[2].Shape.FillRadialGradient);
        Assert.Equal(0.5D, radial.EndX);
        Assert.Equal(0.5D, radial.EndY);
        Assert.Equal(0.4D, radial.EndRadiusX);
        Assert.Equal(0.4D, radial.EndRadiusY);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(2, 10).R > raster.GetPixel(2, 10).B);
        Assert.True(raster.GetPixel(18, 10).B > raster.GetPixel(18, 10).R);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("x2=\"200%\"", exported, StringComparison.Ordinal);
        Assert.Contains("x1=\"-100%\"", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderAppliesSupportedGradientTransformsAndDiagnosesRotatedRadials() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 80 20'><defs>"
            + "<linearGradient id='turn-base' gradientTransform='rotate(90 .5 .5)' x1='0' y1='.5' x2='1' y2='.5'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<linearGradient id='turn' href='#turn-base'/>"
            + "<radialGradient id='spot' gradientTransform='matrix(.5 0 0 1 .5 0)'><stop stop-color='white'/><stop offset='1' stop-color='navy'/></radialGradient>"
            + "<linearGradient id='move' gradientUnits='userSpaceOnUse' gradientTransform='translate(-10 0)' x1='40' y1='0' x2='60' y2='0'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<radialGradient id='skewed' gradientTransform='skewX(20)'><stop stop-color='white'/><stop offset='1' stop-color='black'/></radialGradient>"
            + "</defs><rect width='20' height='20' fill='url(#turn)'/><rect x='20' width='20' height='20' fill='url(#spot)'/>"
            + "<rect x='40' width='20' height='20' fill='url(#move)'/><rect x='60' width='20' height='20' fill='url(#skewed)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        Assert.Equal(4, drawing!.Shapes.Count);
        OfficeLinearGradient turned = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[0].Shape.FillGradient);
        Assert.Equal(0.5D, turned.StartX, 8);
        Assert.Equal(0D, turned.StartY, 8);
        Assert.Equal(0.5D, turned.EndX, 8);
        Assert.Equal(1D, turned.EndY, 8);
        OfficeRadialGradient spot = Assert.IsType<OfficeRadialGradient>(drawing.Shapes[1].Shape.FillRadialGradient);
        Assert.Equal(0.75D, spot.EndX, 8);
        Assert.Equal(0.25D, spot.EndRadiusX, 8);
        Assert.Equal(0.5D, spot.EndRadiusY, 8);
        OfficeLinearGradient moved = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[2].Shape.FillGradient);
        Assert.Equal(-0.5D, moved.StartX, 8);
        Assert.Equal(0.5D, moved.EndX, 8);
        Assert.Null(drawing.Shapes[3].Shape.FillRadialGradient);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(10, 2).R > raster.GetPixel(10, 2).B);
        Assert.True(raster.GetPixel(10, 18).B > raster.GetPixel(10, 18).R);
    }

    [Fact]
    public void SvgReaderMaterializesBoundedLinearRepeatAndReflectPaintServers() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 60 20'><defs>"
            + "<linearGradient id='repeat' spreadMethod='repeat' x2='.25'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient>"
            + "<linearGradient id='reflect' href='#repeat' spreadMethod='reflect'/>"
            + "<radialGradient id='radial-repeat' spreadMethod='repeat'><stop stop-color='white'/><stop offset='1' stop-color='black'/></radialGradient>"
            + "</defs><rect width='20' height='20' fill='url(#repeat)'/><rect x='20' width='20' height='20' fill='url(#reflect)'/>"
            + "<rect x='40' width='20' height='20' fill='url(#radial-repeat)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(1, unsupported);
        Assert.Equal(3, drawing!.Shapes.Count);
        OfficeLinearGradient repeat = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[0].Shape.FillGradient);
        OfficeLinearGradient reflect = Assert.IsType<OfficeLinearGradient>(drawing.Shapes[1].Shape.FillGradient);
        Assert.True(repeat.Stops.Count > 4);
        Assert.True(reflect.Stops.Count > 4);
        Assert.Null(drawing.Shapes[2].Shape.FillRadialGradient);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(1, 10).R > raster.GetPixel(1, 10).B);
        Assert.True(raster.GetPixel(4, 10).B > raster.GetPixel(4, 10).R);
        Assert.True(raster.GetPixel(6, 10).R > raster.GetPixel(6, 10).B);
        Assert.True(raster.GetPixel(21, 10).R > raster.GetPixel(21, 10).B);
        Assert.True(raster.GetPixel(26, 10).B > raster.GetPixel(26, 10).R);
        Assert.True(raster.GetPixel(31, 10).R > raster.GetPixel(31, 10).B);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.DoesNotContain("spreadMethod", exported, StringComparison.Ordinal);
        Assert.True(exported.Split(new[] { "<stop " }, StringSplitOptions.None).Length > 10);
    }

    [Fact]
    public void SvgReaderResolvesCurrentColorInGradientDefinitionTree() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 20 10' color='purple'><defs style='color:orange'>"
            + "<linearGradient id='paint' style='color:lime'><stop stop-color='currentColor'/><stop offset='1' color='red' style='color:blue;stop-color:currentColor'/></linearGradient>"
            + "</defs><rect width='20' height='10' fill='url(#paint)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeLinearGradient gradient = Assert.IsType<OfficeLinearGradient>(Assert.Single(drawing!.Shapes).Shape.FillGradient);
        Assert.Equal(OfficeColor.Lime, gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.Blue, gradient.Stops[gradient.Stops.Count - 1].Color);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(raster.GetPixel(2, 5).G > raster.GetPixel(2, 5).B);
        Assert.True(raster.GetPixel(18, 5).B > raster.GetPixel(18, 5).G);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing);
        Assert.Contains("stop-color=\"#00FF00\"", exported, StringComparison.Ordinal);
        Assert.Contains("stop-color=\"#0000FF\"", exported, StringComparison.Ordinal);
    }

    [Fact]
    public void SvgReaderResolvesRgbaCurrentColorInGradientStops() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>"
            + "<defs><linearGradient id='paint'>"
            + "<stop style='color:rgba(36,87,166,0.502);stop-color:currentColor'/>"
            + "<stop offset='1' color='rgb(10%,20%,30%)' stop-color='currentColor'/></linearGradient></defs>"
            + "<rect width='10' height='10' fill='url(#paint)'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeLinearGradient gradient = Assert.IsType<OfficeLinearGradient>(Assert.Single(drawing!.Shapes).Shape.FillGradient);
        Assert.Equal(OfficeColor.FromRgba(36, 87, 166, 128), gradient.Stops[0].Color);
        Assert.Equal(OfficeColor.FromRgb(26, 51, 76), gradient.Stops[gradient.Stops.Count - 1].Color);
    }

    [Fact]
    public void SvgReaderParsesModernSpaceSeparatedRgbColors() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>"
            + "<rect width='10' height='10' fill='rgb(36 87 166)' stroke='rgb(10% 20% 30% / 50%)' stroke-width='1'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        OfficeShape shape = Assert.Single(drawing!.Shapes).Shape;
        Assert.Equal(OfficeColor.FromRgb(36, 87, 166), shape.FillColor);
        Assert.Equal(OfficeColor.FromRgba(26, 51, 76, 128), shape.StrokeColor);
    }

    [Fact]
    public void SvgReaderPreservesRgbaAlphaWhenExportingShapes() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 10 10'>"
            + "<rect width='10' height='10' fill='rgba(36,87,166,0.502)' stroke='rgba(10,20,30,0.25)' stroke-width='1'/></svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);

        string exported = OfficeDrawingSvgExporter.ToSvg(drawing!);

        Assert.Contains("fill-opacity=\"0.502\"", exported, StringComparison.Ordinal);
        Assert.Contains("stroke-opacity=\"0.251\"", exported, StringComparison.Ordinal);
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
