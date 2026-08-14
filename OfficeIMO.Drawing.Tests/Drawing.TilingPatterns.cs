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
    public void OfficeDrawingTilingPattern_PreservesNearestNeighborImagesThroughTransform() {
        var source = new OfficeRasterImage(2, 1, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        var tile = new OfficeDrawing(2D, 1D);
        tile.AddImageWithInterpolation(
            OfficePngWriter.Encode(source),
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);
        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 4D, 1D),
            2D,
            1D,
            repeatX: false,
            repeatY: false,
            transform: OfficeTransform.Scale(2D, 1D));

        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);

        Assert.Equal(OfficeColor.Black, raster.GetPixel(1, 0));
        Assert.Equal(OfficeColor.White, raster.GetPixel(2, 0));
    }

    [Fact]
    public void OfficeDrawingTilingPattern_IgnoresOffCanvasNearestNeighborImagesForSampling() {
        var source = new OfficeRasterImage(2, 1, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        byte[] png = OfficePngWriter.Encode(source);
        var tile = new OfficeDrawing(2D, 1D);
        tile.AddImageWithInterpolation(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: true);
        var hidden = new OfficeDrawing(2D, 1D);
        hidden.AddImageWithInterpolation(
            png,
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: false);
        tile.AddEffectDrawing(hidden, OfficeTransform.Translate(10D, 0D));
        var drawing = new OfficeDrawing(4D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 4D, 1D),
            2D,
            1D,
            repeatX: false,
            repeatY: false,
            transform: OfficeTransform.Scale(2D, 1D));

        OfficeColor boundary = OfficeDrawingRasterRenderer.Render(drawing).GetPixel(1, 0);

        Assert.InRange(boundary.R, (byte)1, (byte)254);
        Assert.Equal(boundary.R, boundary.G);
        Assert.Equal(boundary.R, boundary.B);
    }

    [Fact]
    public void OfficeDrawingTilingPattern_BoundsScaledIntermediateRaster() {
        var tile = new OfficeDrawing(4000D, 4000D);
        var drawing = new OfficeDrawing(1D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 1D, 1D),
            4000D,
            4000D,
            repeatX: false,
            repeatY: false);

        OfficeImageExportLimitException exception = Assert.Throws<OfficeImageExportLimitException>(
            () => OfficeDrawingRasterRenderer.Render(drawing, scale: 2D));

        Assert.Equal(64_000_000L, exception.RequestedPixels);
        Assert.Equal(OfficeImageExportOptions.DefaultMaximumRasterPixels, exception.MaximumPixels);
    }

    [Fact]
    public void OfficeDrawingTilingPattern_HonorsCallerRasterCeilingForIntermediateTile() {
        var tile = new OfficeDrawing(1000D, 1000D);
        var drawing = new OfficeDrawing(1D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 1D, 1D),
            1000D,
            1000D,
            repeatX: false,
            repeatY: false);

        OfficeImageExportLimitException exception = Assert.Throws<OfficeImageExportLimitException>(
            () => OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
                MaximumRasterPixels = 500_000L
            }));

        Assert.Equal(1_000_000L, exception.RequestedPixels);
        Assert.Equal(500_000L, exception.MaximumPixels);
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
    public void OfficeDrawingSvgExporter_BoundsExpansionAcrossSiblingPatterns() {
        var tile = new OfficeDrawing(1D, 1D);
        OfficeShape square = OfficeShape.Rectangle(1D, 1D);
        square.FillColor = OfficeColor.Red;
        square.StrokeWidth = 0D;
        tile.AddShape(square, 0D, 0D);
        var drawing = new OfficeDrawing(2D, 2D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 2D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false,
            maximumTileCount: 2);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 1D, 2D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false,
            originY: 1D,
            maximumTileCount: 2);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
            () => OfficeDrawingSvgExporter.ToSvg(drawing));

        Assert.Contains("aggregate expansion", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_DefinesTilePayloadOnceAndReusesPlacements() {
        var source = new OfficeRasterImage(2, 1, OfficeColor.Black);
        source.SetPixel(1, 0, OfficeColor.White);
        var tile = new OfficeDrawing(2D, 1D);
        tile.AddImageWithInterpolation(
            OfficePngWriter.Encode(source),
            "image/png",
            new OfficeImageProjection(new OfficeImagePlacement(0D, 0D, 2D, 1D)),
            interpolate: true);
        var drawing = new OfficeDrawing(6D, 1D);
        drawing.AddTilingPattern(
            tile,
            new OfficeImagePlacement(0D, 0D, 6D, 1D),
            2D,
            1D,
            repeatX: true,
            repeatY: false);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(1, CountOccurrences(svg, "data:image/png;base64,"));
        Assert.Equal(3, CountOccurrences(svg, "<use href=\"#officeimo-pattern-tile-"));
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

    [Fact]
    public void OfficeDrawingSvgExporter_SkipsTransparentEffectGroupBeforeExpansionBudget() {
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
        var hidden = new OfficeDrawing(129D, 1D);
        hidden.AddTilingPattern(
            nestedTile,
            new OfficeImagePlacement(0D, 0D, 129D, 1D),
            1D,
            1D,
            repeatX: true,
            repeatY: false);
        var drawing = new OfficeDrawing(129D, 1D);
        drawing.AddEffectDrawing(hidden, OfficeTransform.Identity, 0D);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.DoesNotContain("officeimo-pattern-clip-", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("opacity=\"0\"", svg, StringComparison.Ordinal);
    }
}
