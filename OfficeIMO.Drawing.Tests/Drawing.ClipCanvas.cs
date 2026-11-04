using OfficeIMO.Drawing;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void NonAntialiasedPictureMasksChargeEdgeChecksBeforeExcessivePaintWork() {
        var commands = new List<OfficePathCommand> { OfficePathCommand.MoveTo(0, 0) };
        for (int i = 0; i < 32; i++) {
            commands.Add(OfficePathCommand.LineTo(20, 0));
            commands.Add(OfficePathCommand.LineTo(20, 20));
            commands.Add(OfficePathCommand.LineTo(0, 20));
            commands.Add(OfficePathCommand.LineTo(0, 0));
        }
        var shape = OfficeShape.Rectangle(20, 20);
        shape.FillColor = OfficeColor.Red;
        var inner = new OfficeDrawing(20, 20).AddShape(shape, 0, 0);
        var drawing = new OfficeDrawing(20, 20).AddClippedDrawing(inner, 0, 0,
            OfficeClipPath.Path(20, 20, commands, OfficeFillRule.NonZero));

        Assert.Throws<InvalidDataException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumClipEdgeChecks = 1_000 }));
        OfficeRasterImage accepted = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumClipEdgeChecks = 100_000 });
        Assert.Equal(OfficeColor.Red, accepted.GetPixel(10, 10));
    }

    [Theory]
    [InlineData(OfficeStrokeLineCap.Round)]
    [InlineData(OfficeStrokeLineCap.Square)]
    public void ClipPath_MoveOnlyClosuresDoNotAcquireSvgStrokeCaps(OfficeStrokeLineCap cap) {
        string Svg(params OfficePathCommand[] commands) {
            var shape = OfficeShape.Path(20, 20, commands);
            shape.StrokeColor = OfficeColor.Blue; shape.StrokeWidth = 8; shape.StrokeLineCap = cap;
            return OfficeDrawingSvgExporter.ToSvg(new OfficeDrawing(20, 20).AddShape(shape, 0, 0));
        }
        string empty = Svg(OfficePathCommand.MoveTo(10, 10), OfficePathCommand.Close());
        Assert.Equal("M10 10Z", OfficeSvgFormatting.FormatPathData(new[] { OfficePathCommand.MoveTo(10, 10), OfficePathCommand.Close() }));
        Assert.Contains("d=\"M10 10\"", empty);
        Assert.DoesNotContain("M10 10Z", empty);
        string drawn = Svg(OfficePathCommand.MoveTo(10, 10), OfficePathCommand.LineTo(10, 10), OfficePathCommand.Close());
        Assert.Contains("M10 10L10 10Z", drawn);
        string mixed = Svg(OfficePathCommand.MoveTo(5, 5), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(10, 10), OfficePathCommand.LineTo(15, 15), OfficePathCommand.Close());
        Assert.Contains("M5 5M10 10L15 15Z", mixed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClipPath_DeclaredCanvasMoveOnlyPathsNeitherPaintNorRevealClippedContent(bool close) {
        var commands = close ? new[] { OfficePathCommand.MoveTo(5, 5), OfficePathCommand.Close() }
            : new[] { OfficePathCommand.MoveTo(5, 5) };
        OfficeShape shape = OfficeShape.Path(20, 20, commands);
        shape.FillColor = OfficeColor.Red; shape.StrokeColor = OfficeColor.Blue; shape.StrokeWidth = 4;
        OfficeRasterImage shapeRaster = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(20, 20).AddShape(shape, 0, 0));
        OfficeClipPath clip = OfficeClipPath.Path(20, 20, commands);
        var content = OfficeShape.Rectangle(20, 20); content.FillColor = OfficeColor.Red;
        var inner = new OfficeDrawing(20, 20).AddShape(content, 0, 0);
        OfficeRasterImage clippedRaster = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(20, 20).AddClippedDrawing(inner, 0, 0, clip));
        Assert.Equal(commands, shape.Clone().PathCommands);
        Assert.Equal(commands, clip.Clone().Commands);
        Assert.All(Enumerable.Range(0, 20), y => Assert.All(Enumerable.Range(0, 20), x => {
            Assert.Equal(0, shapeRaster.GetPixel(x, y).A);
            Assert.Equal(0, clippedRaster.GetPixel(x, y).A);
        }));
    }

    [Fact]
    public void ClipPath_ContentSizedMoveOnlyPathsStillCannotInventACanvas() {
        var commands = new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.MoveTo(20, 20) };
        Assert.Throws<ArgumentException>(() => OfficeShape.Path(commands));
        Assert.Throws<ArgumentException>(() => OfficeClipPath.Path(commands));
    }

    [Fact]
    public void ClipPath_DeclaredCanvasStillRejectsMissingCommands() {
        Assert.Throws<ArgumentException>(() => OfficeShape.Path(20, 20, Array.Empty<OfficePathCommand>()));
        Assert.Throws<ArgumentException>(() => OfficeClipPath.Path(20, 20, Array.Empty<OfficePathCommand>()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClipPath_ZeroAreaDeclaredCanvasDoesNotRevealItsRectangle(bool transformed) {
        var inner = new OfficeDrawing(20, 20).AddShape(new OfficeShape { Kind = OfficeShapeKind.Rectangle, Width = 20, Height = 20, FillColor = OfficeColor.Red }, 0, 0);
        var drawing = new OfficeDrawing(40, 40);
        drawing.AddClippedDrawing(inner, 10, 10, OfficeClipPath.Path(20, 20, new[] {
            OfficePathCommand.MoveTo(5, 5), OfficePathCommand.LineTo(15, 5)
        }), new OfficeImageFrameTransform(transformed ? 30 : 0, 20, 20, false, false));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.All(Enumerable.Range(0, 40), y => Assert.All(Enumerable.Range(0, 40), x => Assert.Equal(0, raster.GetPixel(x, y).A)));
    }


    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClipPath_ZeroAreaShapeClipDoesNotRevealItsRectangle(bool transformed) {
        var shape = OfficeShape.Rectangle(20, 20); shape.FillColor = OfficeColor.Red;
        shape.ClipPath = OfficeClipPath.Path(20, 20, new[] { OfficePathCommand.MoveTo(5, 5), OfficePathCommand.LineTo(15, 5) });
        if (transformed) shape.Transform = OfficeTransform.RotateDegrees(30, 10, 10);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(new OfficeDrawing(40, 40).AddShape(shape, 10, 10));
        Assert.All(Enumerable.Range(0, 40), y => Assert.All(Enumerable.Range(0, 40), x => Assert.Equal(0, raster.GetPixel(x, y).A)));
    }

    [Fact]
    public void ClipPath_DeclaredCanvasKeepsInsetCoordinatesAndLegacyOverloadKeepsContentBounds() {
        var commands = new[] { OfficePathCommand.MoveTo(25, 10), OfficePathCommand.LineTo(75, 10),
            OfficePathCommand.LineTo(25, 40), OfficePathCommand.Close() };
        OfficeClipPath content = OfficeClipPath.Path(commands);
        OfficeClipPath canvas = OfficeClipPath.Path(100, 50, commands, OfficeFillRule.NonZero);
        Assert.Equal(50D, content.Width); Assert.Equal(30D, content.Height);
        Assert.Equal(new OfficePoint(0, 0), content.Commands[0].Point);
        Assert.Equal(100D, canvas.Width); Assert.Equal(50D, canvas.Height);
        Assert.Equal(commands, canvas.Commands);
        Assert.Equal(OfficeFillRule.NonZero, canvas.Clone().FillRule);
        Assert.Equal(commands, canvas.Clone().Commands);
    }

    [Fact]
    public void ClipPath_ScalingKeepsTheDeclaredCanvasAndInsetCoordinates() {
        var commands = new[] { OfficePathCommand.MoveTo(25, 10), OfficePathCommand.LineTo(75, 10),
            OfficePathCommand.LineTo(25, 40), OfficePathCommand.Close() };
        OfficeClipPath scaled = OfficeClipPath.Path(100, 50, commands, OfficeFillRule.NonZero).Scale(2, 3);

        Assert.Equal(200D, scaled.Width); Assert.Equal(150D, scaled.Height);
        Assert.Equal(commands.Select(command => command.Scale(2, 3)), scaled.Commands);
        Assert.Equal(OfficeFillRule.NonZero, scaled.FillRule);
        OfficeClipPath content = OfficeClipPath.Path(commands).Scale(2, 3);
        Assert.Equal(100D, content.Width); Assert.Equal(90D, content.Height);
        Assert.Equal(new OfficePoint(0, 0), content.Commands[0].Point);
    }

    [Fact]
    public void ClipPath_ScalingAHorizontalMaskKeepsItsPositiveCanvas() {
        OfficeClipPath scaled = OfficeClipPath.Path(100, 50, new[] {
            OfficePathCommand.MoveTo(25, 10), OfficePathCommand.LineTo(75, 10)
        }).Scale(2, 3);

        Assert.Equal(200D, scaled.Width); Assert.Equal(150D, scaled.Height);
        Assert.Equal(new OfficePoint(50, 30), scaled.Commands[0].Point);
        Assert.Equal(new OfficePoint(150, 30), scaled.Commands[1].Point);
    }

    [Theory]
    [InlineData(0, 50)]
    [InlineData(100, -1)]
    [InlineData(double.NaN, 50)]
    [InlineData(100, double.PositiveInfinity)]
    public void ClipPath_DeclaredCanvasRequiresFinitePositiveDimensions(double width, double height) {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeClipPath.Path(width, height,
            new[] { OfficePathCommand.MoveTo(0, 0), OfficePathCommand.LineTo(10, 10) }));
    }
}
