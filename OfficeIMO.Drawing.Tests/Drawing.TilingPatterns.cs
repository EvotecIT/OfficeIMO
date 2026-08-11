using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(true, false, 3)]
    [InlineData(false, true, 2)]
    public void OfficeDrawingTilingPattern_PreservesSingleAxisRepetition(bool repeatX, bool repeatY, int expectedCount) {
        var tile = new OfficeDrawing(2D, 2D);
        OfficeShape square = OfficeShape.Rectangle(2D, 2D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);
        var drawing = new OfficeDrawing(6D, 4D);

        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 6D, 4D),
            2D,
            2D,
            repeatX: repeatX,
            repeatY: repeatY);

        OfficeDrawingTilingPattern pattern = Assert.Single(drawing.Elements.OfType<OfficeDrawingTilingPattern>());
        Assert.Equal(repeatX, pattern.RepeatX);
        Assert.Equal(repeatY, pattern.RepeatY);
        Assert.Equal(expectedCount, pattern.GetTileTransforms().Count);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void OfficeDrawingTilingPattern_PreservesRepeatAxesWhenTinted(bool repeatX, bool repeatY) {
        var tile = new OfficeDrawing(2D, 2D);
        OfficeShape square = OfficeShape.Rectangle(2D, 2D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);
        var drawing = new OfficeDrawing(6D, 4D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 6D, 4D),
            2D,
            2D,
            repeatX,
            repeatY);

        drawing.ApplyColorTint(OfficeColor.Blue);

        OfficeDrawingTilingPattern pattern = Assert.Single(drawing.Elements.OfType<OfficeDrawingTilingPattern>());
        Assert.Equal(repeatX, pattern.RepeatX);
        Assert.Equal(repeatY, pattern.RepeatY);
        OfficeDrawingShape shape = Assert.Single(pattern.Tile.Elements.OfType<OfficeDrawingShape>());
        Assert.Equal(OfficeColor.Blue, shape.Shape.FillColor);
    }

    [Fact]
    public void OfficeDrawingTilingPattern_RepeatsVectorContentWithGaps() {
        var tile = new OfficeDrawing(2D, 2D);
        OfficeShape square = OfficeShape.Rectangle(2D, 2D);
        square.FillColor = OfficeColor.Blue;
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);

        var drawing = new OfficeDrawing(10D, 4D);
        drawing.AddTilingPattern(tile, new OfficeImagePlacement(0D, 0D, 10D, 4D), 4D, 4D);

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(OfficeColor.Blue, raster.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(2, 0));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(4, 0));
        Assert.Contains("officeimo-pattern-clip-", svg, StringComparison.Ordinal);
        Assert.Contains("matrix(1 0 0 1 4 0)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingTilingPattern_SupportsOverlappingTransformedTiles() {
        var tile = new OfficeDrawing(4D, 4D);
        OfficeShape square = OfficeShape.Rectangle(4D, 4D);
        square.FillColor = OfficeColor.FromRgba(255, 0, 0, 128);
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);

        var drawing = new OfficeDrawing(12D, 8D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 12D, 8D),
            2D,
            4D,
            OfficeTransform.Translate(1D, 0D));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.InRange(raster.GetPixel(3, 2).A, (byte)190, (byte)193);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_BoundsAggregateNestedTilingExpansion() {
        var leaf = new OfficeDrawing(1D, 1D);
        OfficeShape square = OfficeShape.Rectangle(1D, 1D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        leaf.AddShape(square, 0D, 0D);

        var nestedTile = new OfficeDrawing(129D, 1D);
        nestedTile.AddTilingPattern(
            leaf,
            new OfficeImagePlacement(0D, 0D, 129D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false);
        var drawing = new OfficeDrawing(129D, 1D);
        drawing.AddTilingPattern(
            nestedTile,
            new OfficeImagePlacement(0D, 0D, 129D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => OfficeDrawingSvgExporter.ToSvg(drawing));

        Assert.Contains("aggregate expansion", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_HonorsExplicitTileCountAboveDefaultAggregateLimit() {
        var tile = new OfficeDrawing(1D, 1D);
        OfficeShape square = OfficeShape.Rectangle(1D, 1D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);
        var drawing = new OfficeDrawing(16385D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 16385D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false,
            maximumTileCount: 20000);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("matrix(1 0 0 1 16384 0)", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_SkipsTransparentNestedTilingBeforeExpansionBudget() {
        var leaf = new OfficeDrawing(1D, 1D);
        OfficeShape square = OfficeShape.Rectangle(1D, 1D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        leaf.AddShape(square, 0D, 0D);

        var nestedTile = new OfficeDrawing(129D, 1D);
        nestedTile.AddTilingPattern(
            leaf,
            new OfficeImagePlacement(0D, 0D, 129D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false);
        var drawing = new OfficeDrawing(129D, 1D);
        drawing.AddTilingPattern(
            nestedTile,
            new OfficeImagePlacement(0D, 0D, 129D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false,
            opacity: 0D);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.DoesNotContain("officeimo-pattern-clip-", svg, StringComparison.Ordinal);
    }
}
