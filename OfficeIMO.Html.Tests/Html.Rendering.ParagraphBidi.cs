using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("שלום ", "עולם", "םלוע םולש", false)]
    [InlineData("سلام ", "بكم", "\uFEE2\uFEDC\uFE91 \uFEE1\uFE8E\uFEE0\uFEB3", true)]
    public void HtmlRenderer_OrdersPlainRtlWordsAcrossStyledInlineRuns(
        string firstWord,
        string secondWord,
        string expectedVisualText,
        bool expectedShapingLoss) {
        string html = "<p style=\"margin:0;width:400px;direction:ltr;font-family:'" + ManagedTextShapingTestAssets.FamilyName + "'\">"
            + "<span style='color:#ff0000'>" + firstWord + "</span>"
            + "<span style='color:#0000ff'>" + secondWord + "</span></p>";

        // Pin the original scalar/presentation-form fallback contract. The real-font
        // OpenType positive cases are separate; this font intentionally has no GSUB.
        byte[] fontData = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(
            (firstWord + secondWord + expectedVisualText).Select(character => (int)character).ToArray());
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, AllowSystemFontFallback = false };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, fontData);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] text = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .OrderBy(static run => run.X)
            .ToArray();

        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(fontData)!;
        int[] expectedGlyphs = expectedVisualText.Select(scalar => {
            Assert.True(font.TryGetGlyphMetrics(scalar, out int glyph, out _));
            return glyph;
        }).ToArray();
        PaintedTestGlyph[] painted = ObservePaintedTrueTypeGlyphs(rendered, text);
        Assert.All(painted, glyph => Assert.Equal(font.Fingerprint, glyph.Program.Fingerprint));
        Assert.Equal(expectedGlyphs, painted.Select(glyph => glyph.GlyphId));
        Assert.Contains(firstWord + secondWord, rendered.Text, StringComparison.Ordinal);
        Assert.All(text, static run => Assert.True(run.BidiVisualOrderResolved));
        Assert.Equal(expectedShapingLoss, rendered.Diagnostics.Any(diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation));
        Assert.Contains(text, static run => run.Color == OfficeColor.FromRgb(255, 0, 0));
        Assert.Contains(text, static run => run.Color == OfficeColor.FromRgb(0, 0, 255));
        if (!expectedShapingLoss) rendered.RequireNoLoss();
    }
}
