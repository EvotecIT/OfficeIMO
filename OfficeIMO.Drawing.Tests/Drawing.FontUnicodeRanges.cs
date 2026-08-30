using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingFontUnicodeRangeTests {
    [Fact]
    public void UnicodeRangeSet_ParsesNormalizesAndMatchesCssDescriptors() {
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss(
            "U+0000-007F, U+4??, U+1F600",
            out OfficeFontUnicodeRangeSet? ranges));

        Assert.NotNull(ranges);
        Assert.True(ranges!.Contains('A'));
        Assert.True(ranges.Contains(0x0401));
        Assert.True(ranges.Contains(0x1F600));
        Assert.False(ranges.Contains(0x05D0));
        Assert.Equal(3, ranges.Ranges.Count);

        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+??", out OfficeFontUnicodeRangeSet? wildcard));
        Assert.True(wildcard!.Contains(0x00FF));
        Assert.False(wildcard.Contains(0x0100));
    }

    [Theory]
    [InlineData("")]
    [InlineData("U+")]
    [InlineData("U+110000")]
    [InlineData("U+0100-0001")]
    [InlineData("U+4?0")]
    [InlineData("not-a-range")]
    public void UnicodeRangeSet_RejectsInvalidCssDescriptors(string value) {
        Assert.False(OfficeFontUnicodeRangeSet.TryParseCss(value, out _));
    }

    [Fact]
    public void FontCollection_UsesRangesBeforeGlyphCoverageForOneCssFamily() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 0x05D0);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0000-007F", out OfficeFontUnicodeRangeSet? latin));
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0590-05FF", out OfficeFontUnicodeRangeSet? hebrew));
        var fonts = new OfficeFontFaceCollection();

        Assert.True(fonts.TryAdd("Scoped", font, OfficeFontStyle.Regular, latin));
        Assert.True(fonts.TryAdd("Scoped", font, OfficeFontStyle.Regular, hebrew));

        IReadOnlyList<OfficeFontFallbackRun> runs = fonts.PlanFallbackRuns("A\u05D0", "Scoped");

        Assert.Collection(
            runs,
            run => {
                Assert.Equal("A", run.Text);
                Assert.StartsWith("Scoped__officeimo_", run.FamilyName, StringComparison.Ordinal);
            },
            run => {
                Assert.Equal("\u05D0", run.Text);
                Assert.StartsWith("Scoped__officeimo_", run.FamilyName, StringComparison.Ordinal);
                Assert.NotEqual(runs[0].FamilyName, run.FamilyName);
            });
    }

    [Fact]
    public void FontCollection_IgnoresJoiningControlsWhenMatchingUnicodeRanges() {
        const string text = "A\u200D";
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A');
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0041", out OfficeFontUnicodeRangeSet? latin));
        var fonts = new OfficeFontFaceCollection().Add("Scoped", font, OfficeFontStyle.Regular, latin!);
        string resourceFamily = Assert.Single(fonts.Faces).ResourceFamilyName;

        OfficeFontFallbackRun run = Assert.Single(fonts.PlanFallbackRuns(text, "Scoped"));

        Assert.Equal(text, run.Text);
        Assert.Equal(resourceFamily, run.FamilyName);
    }

    [Fact]
    public void FontCollection_UsesRegularRangeFallbackAfterExactStyleCandidates() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 0x05D0);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0000-007F", out OfficeFontUnicodeRangeSet? latin));
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0590-05FF", out OfficeFontUnicodeRangeSet? hebrew));
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", font, OfficeFontStyle.Bold, latin!)
            .Add("Scoped", font, OfficeFontStyle.Regular, hebrew!);
        string boldFamily = Assert.Single(fonts.Faces, face => face.Style == OfficeFontStyle.Bold).ResourceFamilyName;
        string regularFamily = Assert.Single(fonts.Faces, face => face.Style == OfficeFontStyle.Regular).ResourceFamilyName;

        IReadOnlyList<OfficeFontFallbackRun> runs = fonts.PlanFallbackRuns("A\u05D0", "Scoped", OfficeFontStyle.Bold);

        Assert.Collection(
            runs,
            run => {
                Assert.Equal("A", run.Text);
                Assert.Equal(boldFamily, run.FamilyName);
            },
            run => {
                Assert.Equal("\u05D0", run.Text);
                Assert.Equal(regularFamily, run.FamilyName);
            });
    }

    [Fact]
    public void FontCollection_PreservesExplicitRangeSelectionAfterArabicShaping() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont(
            0x0627,
            0x0628,
            0xFE8D,
            0xFE8F);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss(
            "U+0600-06FF",
            out OfficeFontUnicodeRangeSet? arabic));
        var fonts = new OfficeFontFaceCollection().Add(
            "Scoped",
            font,
            OfficeFontStyle.Regular,
            arabic!);
        OfficeFontFallbackRun selected = Assert.Single(
            fonts.PlanFallbackRuns("\u0627\u0628", "Scoped"));
        string shaped = OfficeArabicTextShaper.Shape(selected.Text);

        IReadOnlyList<OfficeFontFallbackRun> shapedRuns = fonts.PlanFallbackRuns(
            shaped,
            selected.FamilyName);

        OfficeFontFallbackRun retained = Assert.Single(shapedRuns);
        Assert.Equal(selected.FamilyName, retained.FamilyName);
        Assert.True(fonts.TryMeasureText(
            shaped,
            12D,
            selected.FamilyName,
            OfficeFontStyle.Regular,
            out double width));
        Assert.True(width > 0D);
    }

    [Fact]
    public void FontCollection_MeasuresAndRasterizesMixedUnicodeRangeRuns() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 0x05D0);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0000-007F", out OfficeFontUnicodeRangeSet? latin));
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0590-05FF", out OfficeFontUnicodeRangeSet? hebrew));
        var fonts = new OfficeFontFaceCollection()
            .Add("Scoped", font, OfficeFontStyle.Regular, latin!)
            .Add("Scoped", font, OfficeFontStyle.Regular, hebrew!);

        Assert.True(fonts.TryMeasureText("A\u05D0", 12D, "Scoped", OfficeFontStyle.Regular, out double width));
        Assert.True(width > 0D);

        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var image = new OfficeRasterImage(80, 30, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(
            image,
            font: null,
            fonts: fonts,
            textShapingProvider: provider);

        canvas.DrawTextLine(
            "A\u05D0",
            2D,
            2D,
            24D,
            OfficeColor.Black,
            alignment: OfficeTextAlignment.Left,
            fontFamily: "Scoped");

        Assert.Collection(
            provider.Requests,
            request => Assert.Equal("A", request.Text),
            request => Assert.Equal("\u05D0", request.Text));
        Assert.Contains(image.GetPixels(), channel => channel == 0);
    }

    [Fact]
    public void SvgExporter_PreservesSelectionFamilyAndUnicodeRangesForDirectDrawingText() {
        byte[] font = ManagedTextShapingTestAssets.CreateFont('A', 0x05D0);
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0000-007F", out OfficeFontUnicodeRangeSet? latin));
        Assert.True(OfficeFontUnicodeRangeSet.TryParseCss("U+0590-05FF", out OfficeFontUnicodeRangeSet? hebrew));
        var drawing = new OfficeDrawing(120D, 30D);
        drawing.Fonts.Add("Scoped", font, OfficeFontStyle.Regular, latin);
        drawing.Fonts.Add("Scoped", font, OfficeFontStyle.Regular, hebrew);
        drawing.AddText("A\u05D0", 0D, 0D, 120D, 30D, new OfficeFontInfo("Scoped", 12D));

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Equal(2, CountOccurrences(svg, "@font-face{font-family:\"Scoped\""));
        Assert.Contains("unicode-range:U+0-7F", svg, StringComparison.Ordinal);
        Assert.Contains("unicode-range:U+590-5FF", svg, StringComparison.Ordinal);
        Assert.Contains("font-family=\"Scoped\"", svg, StringComparison.Ordinal);
        foreach (OfficeFontFace face in drawing.Fonts.Faces) {
            string declaration = Assert.Single(
                svg.Split(new[] { "@font-face{" }, StringSplitOptions.RemoveEmptyEntries),
                value => value.StartsWith("font-family:\"" + face.ResourceFamilyName + "\"", StringComparison.Ordinal));
            Assert.DoesNotContain("unicode-range:", declaration.Split('}')[0], StringComparison.Ordinal);
        }
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int offset = 0;
        while ((offset = value.IndexOf(token, offset, StringComparison.Ordinal)) >= 0) {
            count++;
            offset += token.Length;
        }
        return count;
    }
}
