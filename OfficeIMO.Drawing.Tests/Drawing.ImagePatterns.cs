using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeImagePatternLayout_EnumeratesOnlyIntersectingTiles() {
        var layout = new OfficeImagePatternLayout(
            new OfficeImagePlacement(0D, 0D, 10D, 10D),
            new OfficeImagePlacement(2D, 3D, 4D, 4D));

        IReadOnlyList<OfficeImagePlacement> tiles = layout.GetTilePlacements(16);

        Assert.Equal(9L, layout.EstimatedTileCount);
        Assert.Equal(9, tiles.Count);
        Assert.Equal((-2D, -1D, 4D, 4D), (tiles[0].X, tiles[0].Y, tiles[0].Width, tiles[0].Height));
        Assert.Equal((6D, 7D, 4D, 4D), (tiles[8].X, tiles[8].Y, tiles[8].Width, tiles[8].Height));
        Assert.Throws<InvalidOperationException>(() => layout.GetTilePlacements(8));
    }

    [Fact]
    public void OfficeImagePatternLayout_UsesIndependentRepeatStepsAcrossTransforms() {
        var layout = new OfficeImagePatternLayout(
            new OfficeImagePlacement(0D, 0D, 30D, 4D),
            new OfficeImagePlacement(0D, 0D, 6D, 4D),
            repeatX: true,
            repeatY: false,
            horizontalStep: 12D);

        IReadOnlyList<OfficeImagePlacement> tiles = layout.GetTilePlacements(8);
        OfficeImagePatternLayout translated = layout.Translate(5D, 3D);
        OfficeImagePatternLayout scaled = layout.Scale(2D);

        Assert.Equal(3L, layout.EstimatedTileCount);
        Assert.Equal(new[] { 0D, 12D, 24D }, tiles.Select(tile => tile.X));
        Assert.All(tiles, tile => Assert.Equal(6D, tile.Width));
        Assert.Equal(12D, translated.HorizontalStep);
        Assert.Equal(24D, scaled.HorizontalStep);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImagePatternLayout(
            new OfficeImagePlacement(0D, 0D, 10D, 10D),
            new OfficeImagePlacement(0D, 0D, 2D, 2D),
            horizontalStep: 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImagePatternLayout(
            new OfficeImagePlacement(0D, 0D, 10D, 10D),
            new OfficeImagePlacement(0D, 0D, 2D, 2D),
            horizontalStep: 1D));
    }

    [Fact]
    public void OfficeDrawingImagePattern_PreservesGapsInRasterAndSvg() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Red));
        var layout = new OfficeImagePatternLayout(
            new OfficeImagePlacement(0D, 0D, 10D, 2D),
            new OfficeImagePlacement(0D, 0D, 2D, 2D),
            repeatX: true,
            repeatY: false,
            horizontalStep: 4D);
        var drawing = new OfficeDrawing(10D, 2D).AddImagePattern(png, "image/png", layout, maximumTileCount: 8);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(OfficeColor.Red, rendered.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Transparent, rendered.GetPixel(2, 0));
        Assert.Equal(OfficeColor.Red, rendered.GetPixel(4, 0));
        Assert.Contains("<pattern", svg, StringComparison.Ordinal);
        Assert.Contains("width=\"4\" height=\"2\"><image", svg, StringComparison.Ordinal);
        Assert.Contains("width=\"2\" height=\"2\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingImagePattern_RendersClippedRasterAndCompactSvg() {
        var source = new OfficeRasterImage(2, 1);
        source.SetPixel(0, 0, OfficeColor.Red);
        source.SetPixel(1, 0, OfficeColor.Blue);
        byte[] png = OfficePngWriter.Encode(source);
        var layout = new OfficeImagePatternLayout(
            new OfficeImagePlacement(1D, 1D, 6D, 3D),
            new OfficeImagePlacement(0D, 1D, 2D, 1D));
        var drawing = new OfficeDrawing(8D, 5D)
            .AddImagePattern(png, "image/png", layout, maximumTileCount: 32);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(OfficeColor.Transparent, rendered.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Blue, rendered.GetPixel(1, 1));
        Assert.Equal(OfficeColor.Red, rendered.GetPixel(2, 1));
        Assert.Equal(OfficeColor.Blue, rendered.GetPixel(5, 3));
        Assert.Equal(OfficeColor.Transparent, rendered.GetPixel(7, 4));
        Assert.Single(drawing.ImagePatterns);
        Assert.Empty(OfficeDrawingQualityAnalyzer.Analyze(drawing).Issues);
        Assert.Contains("<pattern", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"url(#officeimo-image-pattern-", svg, StringComparison.Ordinal);
        Assert.Equal(1, CountPatternOccurrences(svg, "data:image/png;base64,"));
    }

    [Fact]
    public void OfficeDrawingImagePattern_RendersInsideATransformedNestedGroup() {
        var sourceImage = new OfficeRasterImage(1, 1, OfficeColor.Red);
        byte[] png = OfficePngWriter.Encode(sourceImage);
        var child = new OfficeDrawing(8D, 8D)
            .AddImagePattern(
                png,
                "image/png",
                new OfficeImagePatternLayout(
                    new OfficeImagePlacement(0D, 0D, 8D, 8D),
                    new OfficeImagePlacement(0D, 0D, 2D, 2D)),
                maximumTileCount: 32);
        var parent = new OfficeDrawing(20D, 20D);
        parent.AddDrawing(child, 6D, 6D, new OfficeImageFrameTransform(0D, 10D, 10D, flipHorizontal: true));

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(parent);

        Assert.Single(parent.Elements.OfType<OfficeDrawingGroup>());
        Assert.Contains(
            Enumerable.Range(0, rendered.Width).SelectMany(x => Enumerable.Range(0, rendered.Height).Select(y => rendered.GetPixel(x, y))),
            color => color.A > 0);
    }

    private static int CountPatternOccurrences(string value, string marker) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(marker, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += marker.Length;
        }

        return count;
    }
}
