using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlArabicProvider_NeutralBracketsRetainResolvedParagraphDirection(bool paintBoundary) {
        byte[] font = PreparedArabicProviderFont();
        var provider = new PreparedArabicContextProvider(font);
        string markup = (paintBoundary ? PreparedArabicColorWord : ArabicContextWord) + " (123)";
        HtmlRenderDocument rendered = RenderArabicContext(markup, provider: provider, fontData: font);
        HtmlRenderText[] runs = ArabicContextText(rendered).ToArray();
        var glyphs = runs.Where(run => run.PreparedShaping != null)
            .SelectMany(run => run.PreparedShaping!.Result.Glyphs.Select(glyph =>
                (Glyph: glyph, Direction: run.PreparedShaping.Result.Direction))).ToArray();

        Assert.Equal(ArabicContextWord + " (123)", rendered.Text.Trim());
        // The closing source bracket belongs to the resolved RTL paragraph even
        // when its local source slice contains only punctuation and whitespace.
        Assert.Contains(glyphs, item => item.Glyph.UnicodeText == ")"
            && item.Direction == OfficeTextDirection.RightToLeft
            && item.Glyph.GlyphId == provider.GlyphId('('));
        Assert.Contains(glyphs, item => item.Glyph.UnicodeText == "("
            && item.Direction == OfficeTextDirection.RightToLeft
            && item.Glyph.GlyphId == provider.GlyphId(')'));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("<u style='color:red'>", "</u>")]
    [InlineData("<s style='color:red'>", "</s>")]
    [InlineData("<a style='color:red' href='https://example.com/arabic'>", "</a>")]
    public void HtmlArabicProvider_DecorationOnlyBoundaryKeepsJoinedWord(string opening, string closing) {
        byte[] font = PreparedArabicProviderFont();
        HtmlRenderDocument rendered = RenderArabicContext(opening
            + "م" + closing
            + "<span style='color:blue'>رحبا</span>",
            provider: new PreparedArabicContextProvider(font), fontData: font);
        HtmlRenderText[] runs = ArabicContextText(rendered).ToArray();

        Assert.Equal(ArabicContextWord, rendered.Text.Trim());
        Assert.Equal(20D, runs.Sum(run => run.TextAdvanceWidth!.Value), 6);
        Assert.All(runs, run => Assert.NotNull(run.PreparedShaping));
        Assert.NotEmpty(DecorationLines(rendered));
        rendered.RequireNoLoss();
    }
}
