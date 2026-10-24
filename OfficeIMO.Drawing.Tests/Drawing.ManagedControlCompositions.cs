using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Threading;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingManagedControlCompositionTests {
    [Fact]
    public void RetainedControlGlyphsDoNotCreateTrackingOrShiftVisibleContoursAndPreparedSlices() {
        byte[] data = TrackedControls();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Tracked controls", data).Faces).Program;
        const string source = "\u202Eaa\u202C";
        OfficeTextShapingResult shaped = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(source,
                "Tracked controls", data, false, font.UnitsPerEm, OfficeTextDirection.RightToLeft)));

        Assert.Equal(source, string.Concat(shaped.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        Assert.Equal(13.5D, font.Measure("aa", 15D), 9);
        Assert.Equal(13.5D, font.MeasureShapedText(source, shaped, 15D), 9);
        EqualContours(font.GetTextContours("aa", 0D, 0D, 15D), font.GetShapedTextContours(source, shaped, 0D, 0D, 15D));

        var prepared = new OfficePreparedTextShaping(font, source, shaped);
        Assert.Equal(13.5D, prepared.Measure(15D), 9);
        Assert.True(prepared.TrySlice(1, 2, out OfficePreparedTextShaping? visible));
        Assert.Equal(13.5D, visible!.Measure(15D), 9);
        EqualContours(prepared.GetContours(0D, 0D, 15D, 1000, CancellationToken.None),
            visible.GetContours(0D, 0D, 15D, 1000, CancellationToken.None));
        foreach (int start in new[] { 0, 3 }) {
            Assert.True(prepared.TrySlice(start, 1, out OfficePreparedTextShaping? control));
            Assert.Equal(0D, control!.Measure(15D));
            Assert.Empty(control.GetContours(0D, 0D, 15D, 1000, CancellationToken.None));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SignedTrackingCountsVisibleClustersAcrossSourceOnlyControls(bool negative) {
        byte[] data = TrackedControls();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Signed controls", data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics(' ', out int blank, out _));
        Assert.True(font.TryGetGlyphMetrics('a', out int visible, out _));
        const string source = "\u202Eaa\u202C";
        var glyphs = new[] {
            new OfficeShapedGlyph(blank, "\u202C", 3, 0),
            new OfficeShapedGlyph(visible, "a", 2, negative ? -500 : 500),
            new OfficeShapedGlyph(visible, "a", 1, negative ? -500 : 500),
            new OfficeShapedGlyph(blank, "\u202E", 0, 0)
        };
        var result = new OfficeTextShapingResult(glyphs, OfficeTextDirection.RightToLeft);

        Assert.Equal(13.5D, font.MeasureShapedText(source, result, 15D), 9);
        var contours = font.GetShapedTextContours(source, result, 0D, 0D, 15D);
        int perGlyph = font.GetTextContours("a", 0D, 0D, 15D).Count;
        Assert.Equal(negative ? -6.75D : 6.75D,
            contours[perGlyph].Min(point => point.X) - contours[0].Min(point => point.X), 9);
    }

    [Theory]
    [InlineData(" ", 500, 6.75D)]
    [InlineData("\u0301", 0, 0.75D)]
    public void VisibleSpaceAndZeroAdvanceMarkStillOwnTheirGraphemeTracking(string visibleText, int advance, double expected) {
        byte[] data = TrackedControls();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Visible tracking", data).Faces).Program;
        Assert.True(font.TryGetGlyphMetrics(' ', out int blank, out _));
        Assert.True(font.TryGetGlyphMetrics('a', out int mark, out _));
        string source = "\u202E" + visibleText + "\u202C";
        var result = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(blank, "\u202E", 0, 0),
            new OfficeShapedGlyph(visibleText == " " ? blank : mark, visibleText, 1, advance),
            new OfficeShapedGlyph(blank, "\u202C", 2, 0)
        }, OfficeTextDirection.LeftToRight);

        Assert.Equal(expected, font.MeasureShapedText(source, result, 15D), 9);
    }

    [Theory]
    [InlineData('\u200B', false, false, false)]
    [InlineData('\u200B', true, false, false)]
    [InlineData('\u200B', false, true, false)]
    [InlineData('\u200B', true, true, true)]
    [InlineData('\u034F', true, false, false)]
    [InlineData('\u034F', true, true, true)]
    public void EveryFinalControlGlyphMustRemainBlankAfterGsubIncludingContinuations(char control,
        bool blankFirst, bool blankSecond, bool accepted) {
        byte[] data = ManagedTextShapingTestAssets.CreateControlSubstitutionFont(blankFirst, blankSecond);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Control GSUB", data).Faces).Program;
        string source = "A" + control + "A";
        OfficeTextShapingResult? result = OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(source, "Control GSUB", data, false, font.UnitsPerEm,
                OfficeTextDirection.LeftToRight, "en"));

        if (!accepted) { Assert.Null(result); return; }
        Assert.NotNull(result);
        Assert.Equal(source, string.Concat(result!.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        OfficeShapedGlyph[] ghosts = result.Glyphs.Where(glyph => glyph.GlyphId is 3 or 4).ToArray();
        Assert.Equal(2, ghosts.Length);
        Assert.All(ghosts, glyph => { Assert.Equal(0, glyph.AdvanceWidth); Assert.Equal(0, glyph.AdvanceHeight); });
        Assert.Equal(font.GetTextContours("AA", 0D, 0D, 20D).Count,
            font.GetShapedTextContours(source, result, 0D, 0D, 20D).Count);
    }

    [Fact]
    public void DisabledControlSubstitutionRetainsTheInitiallyProvenBlankGlyph() {
        byte[] data = ManagedTextShapingTestAssets.CreateControlSubstitutionFont(false, false);
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Disabled control GSUB", data).Faces).Program;
        const string source = "A\u200BA";
        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(source, "Disabled control GSUB", data, false, font.UnitsPerEm,
                OfficeTextDirection.LeftToRight, "en", OfficeTextFeatureSettings.Default.With("ccmp", 0))));

        Assert.Equal(source, string.Concat(result.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        Assert.Equal(0, Assert.Single(result.Glyphs, glyph => glyph.GlyphId == 1).AdvanceWidth);
    }

    private static byte[] TrackedControls() => ManagedTextShapingTestAssets.AddTrackingTable(
        ManagedTextShapingTestAssets.CreateControlAdmissionFont(0x202E, blankSpace: true),
        ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { -1D, 1D },
            new[] { new short[] { -120, -40 }, new short[] { -80, 40 } }));

    private static void EqualContours(List<List<OfficePoint>> expected, List<List<OfficePoint>> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int contour = 0; contour < expected.Count; contour++) Assert.Equal(expected[contour].ToArray(), actual[contour].ToArray());
    }
}
