using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontTrackingTests {
    private static byte[] Table() => ManagedTextShapingTestAssets.CreateTrackingTable(
        new[] { 10D, 20D }, new[] { -1D, 1D },
        new[] { new short[] { -120, -40 }, new short[] { -80, 40 } });

    private static IOfficeFontProgram Font(bool color = false) =>
        Assert.IsAssignableFrom<IOfficeFontProgram>(OfficeTrueTypeFont.TryLoad(
            ManagedTextShapingTestAssets.CreateTrackingFont(Table(), color)));

    [Theory]
    [InlineData(5, -150)]
    [InlineData(10, -100)]
    [InlineData(15, -50)]
    [InlineData(20, 0)]
    [InlineData(25, 50)]
    public void NormalTrackingInterpolatesAndExtrapolatesSizesAndTracks(double size, double adjustment) {
        IOfficeFontProgram font = Font();
        Assert.True(font.TryGetGlyphMetrics('A', out _, out int nominal));
        Assert.Equal(500, nominal);
        Assert.Equal(2 * (nominal + adjustment) * size / font.UnitsPerEm, font.Measure("AB", size), 9);
        Assert.Equal(font.Measure("AB", size), font.MeasureTextElements(new[] { "A", "B" }, size).Sum(), 9);
        var contours = font.GetTextContours("AB", 0, 0, size);
        int oneCount = font.GetTextContours("A", 0, 0, size).Count;
        double firstX = contours[0].Min(point => point.X);
        double secondX = contours[oneCount].Min(point => point.X);
        Assert.Equal((nominal + adjustment) * size / font.UnitsPerEm, secondX - firstX, 9);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapedTrackingUsesGlyphCountAndSignedDirection(bool negative) {
        IOfficeFontProgram font = Font();
        int advance = negative ? -500 : 500;
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "AB", 0, advanceWidth: advance)
        }, negative ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        Assert.Equal(6.75D, font.MeasureShapedText("AB", run, 15D), 9);
        Assert.NotEmpty(font.GetShapedTextContours("AB", run, 0, 0, 15D));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapedContoursKeepTheSignedTrackedOrigins(bool negative) {
        IOfficeFontProgram font = Font();
        int advance = negative ? -500 : 500;
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: advance),
            new OfficeShapedGlyph(1, "B", 1, advanceWidth: advance)
        }, negative ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        Assert.Equal(13.5D, font.MeasureShapedText("AB", run, 15D), 9);
        var contours = font.GetShapedTextContours("AB", run, 0, 0, 15D);
        int count = font.GetTextContours("A", 0, 0, 15D).Count;
        Assert.Equal(negative ? -6.75D : 6.75D,
            contours[count].Min(point => point.X) - contours[0].Min(point => point.X), 9);
    }

    [Fact]
    public void HorizontalTrackingDoesNotAlterVerticalAdvances() {
        IOfficeFontProgram font = Font();
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 500, advanceHeight: -1000, offsetX: 0, offsetY: 0),
            new OfficeShapedGlyph(1, "B", 1, advanceWidth: 500, advanceHeight: -1000, offsetX: 0, offsetY: 0)
        }, OfficeTextDirection.TopToBottom);
        Assert.Equal(30D, font.MeasureShapedText("AB", run, 15D), 9);
        var contours = font.GetShapedTextContours("AB", run, 0, 0, 15D);
        int count = font.GetTextContours("A", 0, 0, 15D).Count;
        Assert.Equal(15D, contours[count].Min(point => point.Y) - contours[0].Min(point => point.Y), 9);
    }

    [Fact]
    public void ShapedClusterRetainsItsAttachedMarkAndOneTrackingAdvance() {
        IOfficeFontProgram font = Font();
        var run = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 500),
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 0, offsetX: -500, offsetY: 100)
        }, OfficeTextDirection.LeftToRight);
        Assert.Equal(6.75D, font.MeasureShapedText("A", run, 15D), 9);
        var contours = font.GetShapedTextContours("A", run, 0, 0, 15D);
        int count = font.GetTextContours("A", 0, 0, 15D).Count;
        Assert.Equal(contours[0].Min(point => point.X), contours[count].Min(point => point.X), 9);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RasterScaleKeepsLogicalTrackingSize(bool color, bool shaped) {
        byte[] data = ManagedTextShapingTestAssets.CreateTrackingFont(Table(), color);
        var fonts = new OfficeFontFaceCollection().Add("Tracked", data);
        var program = Assert.Single(fonts.Faces).Program;
        double logicalWidth = program.Measure("AB", 15D);
        var actual = new OfficeDrawing(50, 30).AddPositionedText("AB", 0, 0, logicalWidth, 20,
            new OfficeFontInfo("Tracked", 15), OfficeColor.Black, textAdvanceWidth: logicalWidth);
        // Same logical scene drawn as single glyphs protects positions and outlines against export-size tracking.
        var expected = new OfficeDrawing(50, 30)
            .AddPositionedText("A", 0, 0, logicalWidth / 2, 20, new OfficeFontInfo("Tracked", 15), OfficeColor.Black,
                textAdvanceWidth: 7.5D)
            .AddPositionedText("B", logicalWidth / 2, 0, logicalWidth / 2, 20, new OfficeFontInfo("Tracked", 15), OfficeColor.Black,
                textAdvanceWidth: 7.5D);
        if (shaped) actual.TextShapingProvider = new FixedTrackingShaper();
        actual.Fonts.AddRange(fonts);
        expected.Fonts.Add("Tracked", ManagedTextShapingTestAssets.CreateTrackingFont(null, color));
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected, 2D).GetPixels(),
            OfficeDrawingRasterRenderer.Render(actual, 2D).GetPixels());
    }

    [Fact]
    public void OrdinaryRasterTextRetainsLogicalTrackingAtExportScale() {
        var actual = new OfficeDrawing(100, 30).AddText("AB", 0, 0, 80, 30,
            new OfficeFontInfo("Tracked", 15), OfficeColor.Black);
        actual.Fonts.Add("Tracked", ManagedTextShapingTestAssets.CreateTrackingFont(Table()));
        var expected = new OfficeDrawing(100, 30)
            .AddText("A", 0, 0, 80, 30, new OfficeFontInfo("Tracked", 15), OfficeColor.Black)
            .AddText("B", 6.75D, 0, 80, 30, new OfficeFontInfo("Tracked", 15), OfficeColor.Black);
        expected.Fonts.Add("Tracked", ManagedTextShapingTestAssets.CreateTrackingFont(null));
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected, 2D).GetPixels(),
            OfficeDrawingRasterRenderer.Render(actual, 2D).GetPixels());
    }

    private sealed class FixedTrackingShaper : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => new(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 500),
            new OfficeShapedGlyph(1, "B", 1, advanceWidth: 500)
        }, OfficeTextDirection.LeftToRight);
    }

    [Fact]
    public void ColorContoursUseTheMeasuredTrackingAdvance() {
        byte[] table = ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
            new[] { new short[] { 200, 300 } });
        byte[] data = ManagedTextShapingTestAssets.CreateTrackingFont(table, color: true);
        IOfficeFontProgram font = Assert.IsAssignableFrom<IOfficeFontProgram>(OfficeTrueTypeFont.TryLoad(data));
        var fonts = new OfficeFontFaceCollection().Add("Tracked", data);
        var expected = new OfficeDrawing(40, 30)
            .AddPositionedText("A", 0, 0, 15, 20, new OfficeFontInfo("Tracked", 15), OfficeColor.Black, textAdvanceWidth: font.Measure("A", 15))
            .AddPositionedText("B", font.Measure("A", 15), 0, 15, 20, new OfficeFontInfo("Tracked", 15), OfficeColor.Black, textAdvanceWidth: font.Measure("A", 15));
        var actual = new OfficeDrawing(40, 30)
            .AddPositionedText("AB", 0, 0, 30, 20, new OfficeFontInfo("Tracked", 15), OfficeColor.Black, textAdvanceWidth: font.Measure("AB", 15));
        expected.Fonts.AddRange(fonts); actual.Fonts.AddRange(fonts);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected).GetPixels(),
            OfficeDrawingRasterRenderer.Render(actual).GetPixels());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void MalformedTrackingRejectsTheFontWithoutThrowing(int corruption) {
        byte[] table = Table();
        switch (corruption) {
            case 0: table[3] = 1; break; // Unsupported version.
            case 1: table[13] = 65; break; // Track allocation bound.
            case 2: table[16] = 127; break; // Outside table size offset.
            case 3: Array.Copy(table, 20, table, 28, 4); break; // Duplicate tracks.
            case 4: Array.Copy(table, 36, table, 40, 4); break; // Duplicate sizes.
        }
        Assert.Null(OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateTrackingFont(table)));
    }
}
