using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGraphicsTilingPatternTests {
    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, true, true)]
    public void ReusablePatternPreservesGapsAxesAndTransforms(bool repeatX, bool repeatY, bool transformed) {
        OfficeDrawing drawing = CreatePattern(repeatX, repeatY, transformed);
        byte[] pdf = PdfDocument.Create(Options()).Drawing(drawing).ToBytes();
        string content = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/PatternType 1", content, StringComparison.Ordinal);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        // Points stay inside paint or gaps, away from cell and area edges.
        foreach (var point in new[] { (5, 7), (9, 7), (17, 7), (21, 7), (5, 17), (9, 17), (42, 32), (78, 48), (1, 1), (88, 58) }) {
            OfficeColor a = actual.GetPixel(point.Item1, point.Item2), e = expected.GetPixel(point.Item1, point.Item2);
            Assert.InRange(Math.Abs(a.R - e.R), 0, 4);
            Assert.InRange(Math.Abs(a.G - e.G), 0, 4);
            Assert.InRange(Math.Abs(a.B - e.B), 0, 4);
        }
    }

    [Fact]
    public void NestedPatternRetainsCellResourcesAndGroupOpacity() {
        OfficeDrawing tile = CreatePattern(true, true, false);
        OfficeDrawing drawing = new OfficeDrawing(90D, 60D)
            .AddTilingPattern(tile, new OfficeImagePlacement(0D, 0D, 90D, 60D), 90D, 60D, opacity: 0.5D);
        byte[] pdf = PdfDocument.Create(Options()).Drawing(drawing).ToBytes();
        string content = Encoding.ASCII.GetString(pdf);
        Assert.Equal(2, content.Split(new[] { "/PatternType 1" }, StringSplitOptions.None).Length - 1);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        OfficeColor red = raster.GetPixel(5, 7);
        Assert.InRange(red.R, 250, 255);
        Assert.InRange(red.G, 120, 135);
        Assert.InRange(red.B, 120, 135);
        Assert.Equal(OfficeColor.White, raster.GetPixel(1, 1));
    }

    [Fact]
    public void NestedPatternLinksExpandNormallyWithinTheSharedBudget() {
        OfficeDrawing drawing = CreateNestedLinkedPattern(32);
        var document = PdfDocument.Create(Options()).Drawing(drawing);
        byte[] pdf = document.ToBytes();
        Assert.Equal(16, PdfInspector.Inspect(pdf).GetAnnotationsBySubtype("Link").Count);
        // A new serialization owns a fresh ledger rather than retaining prior work.
        Assert.Equal(pdf, document.ToBytes());
    }

    [Fact]
    public void NestedPatternLinksRejectAggregateExpansionBeforeCopyingOuterAnnotations() {
        // Four inner annotations fit the limit. The sixteen outer copies must be
        // rejected by preflight, without constructing a large expansion fixture.
        var document = PdfDocument.Create(Options()).Drawing(CreateNestedLinkedPattern(4));
        var error = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains("pattern annotation aggregate expansion", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PatternAnnotationBudgetCountsEveryLinkInEachTile() {
        var tile = new OfficeDrawing(3D, 1D)
            .AddLink("https://example.com/first", 0D, 0D, 1D, 1D)
            .AddLink("https://example.com/second", 1D, 0D, 1D, 1D)
            .AddLink("https://example.com/third", 2D, 0D, 1D, 1D);
        var drawing = new OfficeDrawing(6D, 1D).AddTilingPattern(tile,
            new OfficeImagePlacement(0D, 0D, 6D, 1D), 3D, 1D, maximumTileCount: 4);
        var error = Assert.Throws<InvalidOperationException>(() => PdfDocument.Create(Options()).Drawing(drawing).ToBytes());
        Assert.Contains("pattern annotation aggregate expansion", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PatternAnnotationBudgetIsSharedAcrossPagesAndRunningContexts(bool header) {
        var tile = new OfficeDrawing(1D, 1D).AddLink("https://example.com/", 0D, 0D, 1D, 1D);
        var drawing = new OfficeDrawing(2D, 1D).AddTilingPattern(tile,
            new OfficeImagePlacement(0D, 0D, 2D, 1D), 1D, 1D, maximumTileCount: 3);
        var document = PdfDocument.Create(Options());
        if (header) document.Header(builder => builder.Content(content => content.Drawing(drawing, spacingBefore: 0D, spacingAfter: 0D), 8D));
        document.Drawing(drawing);
        if (!header) document.PageBreak().Drawing(drawing);
        var error = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains("pattern annotation aggregate expansion", error.Message, StringComparison.Ordinal);
    }

    private static OfficeDrawing CreateNestedLinkedPattern(int maximumTileCount) {
        var unit = new OfficeDrawing(1D, 1D).AddLink("https://example.com/", 0D, 0D, 1D, 1D);
        var tile = new OfficeDrawing(2D, 2D).AddTilingPattern(unit,
            new OfficeImagePlacement(0D, 0D, 2D, 2D), 1D, 1D, maximumTileCount: maximumTileCount);
        return new OfficeDrawing(4D, 4D).AddTilingPattern(tile,
            new OfficeImagePlacement(0D, 0D, 4D, 4D), 2D, 2D, maximumTileCount: maximumTileCount);
    }

    private static OfficeDrawing CreatePattern(bool repeatX, bool repeatY, bool transformed) {
        var red = OfficeShape.Rectangle(4D, 6D); red.FillColor = OfficeColor.Red; red.StrokeWidth = 0D;
        var blue = OfficeShape.Rectangle(4D, 6D); blue.FillColor = OfficeColor.Blue; blue.StrokeWidth = 0D;
        var tile = new OfficeDrawing(8D, 6D).AddShape(red, 0D, 0D).AddShape(blue, 4D, 0D);
        OfficeTransform transform = transformed ? new OfficeTransform(1D, 0D, 0.25D, 1D, 2D, 1D) : OfficeTransform.Identity;
        return new OfficeDrawing(90D, 60D).AddTilingPattern(tile, new OfficeImagePlacement(3D, 5D, 80D, 50D),
            12D, 10D, repeatX, repeatY, transform, originX: 3D, originY: 5D);
    }

    private static PdfOptions Options() => new PdfOptions { PageWidth = 90D, PageHeight = 60D,
        MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D };
}
