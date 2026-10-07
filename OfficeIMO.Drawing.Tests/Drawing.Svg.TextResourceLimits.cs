using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgTextResourceLimitTests {
    [Fact]
    public void RepeatedTransformedCffRunsRetainTheSameRandomDependentGlyphPaint() {
        byte[] font = CffRandomGlyphTestAssets.CreateRandomOverhangFont();
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("Random CFF", font);
        OfficeDrawing Read(string runs) {
            string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='900' height='120' font-family='Random CFF' font-size='20'>" + runs + "</svg>";
            Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
            Assert.Equal(0, unsupported);
            return drawing!;
        }
        OfficeRasterImage single = OfficeDrawingRasterRenderer.Render(Read("<text transform='translate(100 80)' y='0'>A</text>"));
        OfficeRasterImage expected = new OfficeRasterImage(900, 120);
        int painted = 0;
        for (int y = 0; y < single.Height; y++) {
            for (int x = 0; x < single.Width; x++) {
                OfficeColor pixel = single.GetPixel(x, y);
                if (pixel.A == 0) continue;
                painted++;
                for (int index = 0; index < 8; index++) {
                    if (x + index * 100 < expected.Width) expected.SetPixel(x + index * 100, y, pixel);
                }
            }
        }
        Assert.True(painted > 0);
        string repeated = string.Concat(Enumerable.Range(1, 8).Select(index =>
            "<text transform='translate(" + index * 100 + " 80)' y='0'>A</text>"));
        byte[] actual = OfficeDrawingRasterRenderer.Render(Read(repeated)).GetPixels();
        int maximumDifference = expected.GetPixels().Zip(actual, (left, right) => Math.Abs(left - right)).Max();
        Assert.True(maximumDifference == 0, "Repeated runs changed or clipped glyph paint; maximum channel difference: " + maximumDifference);
    }

    [Theory]
    [InlineData("root")]
    [InlineData("nested")]
    [InlineData("symbol")]
    public void FittedViewportsChargeExpansionOfEveryRetainedEffect(string kind) {
        string groups = string.Concat(Enumerable.Repeat("<g mix-blend-mode='multiply'><rect x='-1' y='1' width='2' height='2'/></g>", 10));
        string content = kind == "root" ? groups : kind == "nested"
            ? "<svg width='4096' height='2048' viewBox='0 0 2048 2048'>" + groups + "</svg>"
            : "<defs><symbol id='s' viewBox='0 0 2048 2048'>" + groups + "</symbol></defs><use href='#s' width='4096' height='2048'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='4096' height='2048' "
            + (kind == "root" ? "viewBox='0 0 2048 2048'" : "") + ">" + content + "</svg>";

        bool success = OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported);
        if (kind == "root") Assert.False(success);
        else {
            Assert.True(success);
            Assert.True(unsupported > 0);
            Assert.True(EffectPixels(drawing!) <= 64_000_000D);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FittedViewportsKeepCompleteSmallTextWithinTheAggregateSurfaceBudget(bool nested) {
        string runs = string.Concat(Enumerable.Range(0, 16).Select(index =>
            "<g transform='translate(" + (index == 0 ? -1 : 100 + index * 30) + " 30)'><text y='0'>A</text></g>"));
        string viewport = "width='4096' height='2048' viewBox='0 0 2048 2048'";
        string content = nested ? "<svg " + viewport + ">" + runs + "</svg>" : runs;
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' " + (nested ? "width='4096' height='2048'" : viewport)
            + " font-family='FixtureFont' font-size='20'>" + content + "</svg>";
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("FixtureFont", ManagedTextShapingTestAssets.CreateFont('A'));

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.Equal(16, TextCount(drawing!));
        Assert.True(EffectPixels(drawing!) <= 64_000_000D, "Retained effects require " + EffectPixels(drawing!) + " pixels.");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TransformedCffMeasurementsShareTheDocumentBudgetAndOmitExhaustedRuns(bool singleRun) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        var options = new OfficeSvgDrawingReaderOptions { MaximumViewportDimension = 100_000D };
        options.Fonts.Add("Source Sans", font);
        OfficeFontFace face = Assert.Single(options.Fonts.Faces);
        var cff = Assert.IsAssignableFrom<IOfficeCffBoundedFontProgram>(face.Program);
        var budget = new OfficeCffOperationBudget();
        Assert.NotEmpty(cff.GetTextContoursBounded("A", 0D, 0D, 1D, 100_000, CancellationToken.None, budget));
        int oneGlyphOperations = 1_000_000 - budget.RemainingOperations;
        Assert.True(oneGlyphOperations > 0);
        int glyphs = 1_000_000 / oneGlyphOperations + 1;
        int runLength = singleRun ? glyphs : 64;
        int runCount = singleRun ? 1 : glyphs / runLength + 1;
        string text = new string('A', runLength);
        string content = string.Concat(Enumerable.Repeat("<text transform='translate(2 2)' y='1'>" + text + "</text>", runCount));
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40' font-family='Source Sans' font-size='1'>" + content + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options, out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.True(unsupported > 0, "All runs were admitted after more than one million CFF operations.");
        if (singleRun) Assert.Equal(0, TextCount(drawing!));
        else Assert.InRange(TextCount(drawing!), 1, runCount - 1);
    }

    private static int TextCount(OfficeDrawing drawing) => drawing.Elements.Sum(element =>
        element is OfficeDrawingText ? 1 : element is OfficeDrawingEffectGroup effect ? TextCount(effect.Drawing) :
        element is OfficeDrawingGroup group ? TextCount(group.Drawing) : 0);

    private static double EffectPixels(OfficeDrawing drawing) => drawing.Elements.Sum(element =>
        element is OfficeDrawingEffectGroup effect ? effect.Drawing.Width * effect.Drawing.Height + EffectPixels(effect.Drawing) :
        element is OfficeDrawingGroup group ? EffectPixels(group.Drawing) : 0D);
}
