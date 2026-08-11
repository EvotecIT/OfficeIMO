using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_HonorsNoWrapAndEndEllipsisInSceneAndPdf() {
        const string html = "<p style='margin:0;width:90px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:14px'>Alpha beta gamma delta</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText text = Assert.Single(EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>());

        Assert.EndsWith("\u2026", text.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("delta", text.Text, StringComparison.Ordinal);
        Assert.InRange(text.TextAdvanceWidth ?? text.Width, 0.01D, 90D);
        string pdfText = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdf(new HtmlPdfSaveOptions())).ExtractText();
        Assert.Contains("\u2026", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("delta", pdfText, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("line-clamp")]
    [InlineData("-webkit-line-clamp")]
    public void HtmlRender_ClampsLinesAndAddsAnEllipsis(string property) {
        string html = "<p style='margin:0;width:100px;overflow:hidden;" + property + ":2;font-size:14px;line-height:18px'>one two three four five six seven eight nine ten eleven twelve</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        IReadOnlyList<HtmlRenderText> lines = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToList();

        Assert.Equal(2, lines.Select(line => line.Y).Distinct().Count());
        Assert.EndsWith("\u2026", lines[lines.Count - 1].Text, StringComparison.Ordinal);
        Assert.InRange(lines.Max(line => line.Y + line.Height) - lines.Min(line => line.Y), 35.9D, 36.1D);
    }

    [Fact]
    public void HtmlRender_UsesInheritedNumericTabStopsForPreformattedText() {
        const string compact = "<div style='tab-size:2'><pre style='margin:0;font-family:Consolas;font-size:12px'>A\tB</pre></div>";
        const string wide = "<div style='tab-size:8'><pre style='margin:0;font-family:Consolas;font-size:12px'>A\tB</pre></div>";

        HtmlRenderText compactText = Assert.Single(EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(compact).Pages[0].Scene).OfType<HtmlRenderText>());
        HtmlRenderText wideText = Assert.Single(EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(wide).Pages[0].Scene).OfType<HtmlRenderText>());

        Assert.DoesNotContain('\t', compactText.Text);
        Assert.DoesNotContain('\t', wideText.Text);
        Assert.True(wideText.TextAdvanceWidth > compactText.TextAdvanceWidth);
        Assert.Contains("B", PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(wide).ToPdf(new HtmlPdfSaveOptions())).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_PositionsLetterAndWordSpacingWithoutCollapsingSearchableSpaces() {
        const string normal = "<p style='margin:0;font-family:Consolas;font-size:14px'>A B</p>";
        const string spaced = "<p style='margin:0;font-family:Consolas;font-size:14px;letter-spacing:2px;word-spacing:6px'>A B</p>";

        HtmlRenderDocument normalRender = HtmlRenderTestDriver.Render(normal);
        HtmlRenderDocument spacedRender = HtmlRenderTestDriver.Render(spaced);
        double normalAdvance = EnumerateTextOverflowVisuals(normalRender.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Sum(text => text.TextAdvanceWidth ?? text.Width);
        IReadOnlyList<HtmlRenderText> spacedGlyphs = EnumerateTextOverflowVisuals(spacedRender.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();

        Assert.Equal(new[] { "A", " ", "B" }, spacedGlyphs.Select(glyph => glyph.Text).ToArray());
        Assert.True(spacedGlyphs.Sum(glyph => glyph.TextAdvanceWidth ?? glyph.Width) > normalAdvance + 10D);
        Assert.True(spacedGlyphs[2].X - spacedGlyphs[0].X > normalAdvance / 2D);
        string pdfText = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(spaced).ToPdf(new HtmlPdfSaveOptions())).ExtractText();
        Assert.Contains("A B", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_LetterSpacingPositionsEveryGraphemeWithoutScalingWholeWords() {
        const string html = "<p style='margin:0;font-family:Consolas;font-size:14px;letter-spacing:3px'>AB</p>";

        IReadOnlyList<HtmlRenderText> glyphs = EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();

        Assert.Equal(new[] { "A", "B" }, glyphs.Select(glyph => glyph.Text).ToArray());
        Assert.True(glyphs[1].X > glyphs[0].X + 3D);
        Assert.All(glyphs, glyph => Assert.True((glyph.TextAdvanceWidth ?? glyph.Width) < 30D));
    }

    [Fact]
    public void HtmlRender_LetterSpacingRetainsPerGraphemePaintForRightToLeftText() {
        const string html = "<p dir='rtl' style='margin:0;font-family:Arial;font-size:14px;letter-spacing:3px'>אב</p>";

        IReadOnlyList<HtmlRenderText> glyphs = EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(html).Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .ToList();

        Assert.Equal(2, glyphs.Count);
        Assert.Equal(new[] { "א", "ב" }, glyphs.Select(glyph => glyph.Text).OrderBy(text => text, StringComparer.Ordinal).ToArray());
        Assert.NotEqual(glyphs[0].X, glyphs[1].X);
        Assert.All(glyphs, glyph => Assert.True((glyph.TextAdvanceWidth ?? glyph.Width) < 30D));
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateTextOverflowVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateTextOverflowVisuals(children)) yield return child;
        }
    }
}
