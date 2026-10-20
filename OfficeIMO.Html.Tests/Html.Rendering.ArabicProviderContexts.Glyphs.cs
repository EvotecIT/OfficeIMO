using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlArabicProvider_RetainsContextGlyphOffsetsWithoutLeakingToIdenticalSourceElsewhere() {
        byte[] font = PreparedArabicProviderFont();
        var provider = new PreparedArabicContextProvider(font);
        HtmlRenderDocument rendered = RenderArabicContext(PreparedArabicColorWord + "<br>"
            + "<span style='color:red'>م</span> <span style='color:blue'>رحبا</span>",
            provider: provider, fontData: font);
        HtmlRenderText[] text = ArabicContextText(rendered).ToArray();
        HtmlRenderText[] red = text.Where(item => item.Color == OfficeColor.Red).OrderBy(item => item.Y).ToArray();
        Assert.Equal(2, red.Length);
        var prepared = red[0].PreparedShaping;
        Assert.NotNull(prepared);
        Assert.Equal("م", prepared!.Text);
        Assert.Equal(provider.Program.Fingerprint, prepared.Program.Fingerprint);
        OfficeShapedGlyph meem = Assert.Single(prepared.Result.Glyphs);
        Assert.Equal(provider.GlyphId(ArabicContextForms[0]), meem.GlyphId);
        Assert.Equal("م", meem.UnicodeText);
        Assert.Equal(0, meem.TextIndex);
        Assert.Equal(160, meem.AdvanceWidth);
        Assert.Equal(11, meem.OffsetX);
        Assert.Equal(-7, meem.OffsetY);
        Assert.Equal(prepared.Measure(20D), red[0].TextAdvanceWidth!.Value, 6);
        Assert.Equal(20D * 700D / provider.Program.UnitsPerEm, red[1].TextAdvanceWidth!.Value, 6);
        HtmlRenderText blue = Assert.Single(text, item => item.Color == OfficeColor.Blue && item.Y == red[0].Y);
        Assert.NotNull(blue.PreparedShaping);
        Assert.Equal("رحبا", blue.PreparedShaping!.Text);
        Assert.Equal(ArabicContextForms.Substring(1).Reverse().Select(provider.GlyphId),
            blue.PreparedShaping.Result.Glyphs.Select(glyph => glyph.GlyphId));
        Assert.Equal(ArabicContextWord + "\nم رحبا", rendered.Text.Trim().Replace("\r\n", "\n"));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlArabicProvider_ClusterAcrossPaintBoundaryReportsLossButWholeClusterDoesNot() {
        byte[] font = PreparedArabicProviderFont();
        HtmlRenderDocument plain = RenderArabicContext(ArabicContextWord,
            provider: new PreparedArabicContextProvider(font, mergeFirstCluster: true), fontData: font);
        plain.RequireNoLoss();
        var provider = new PreparedArabicContextProvider(font, mergeFirstCluster: true);
        HtmlRenderDocument split = RenderArabicContext(PreparedArabicColorWord, provider: provider, fontData: font);

        Assert.Contains(provider.Requests, request => request.Text.Contains(ArabicContextWord, StringComparison.Ordinal));
        Assert.Equal(ArabicContextWord, split.Text.Trim());
        Assert.Contains(split.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<HtmlConversionException>(() => split.RequireNoLoss());
    }
}
