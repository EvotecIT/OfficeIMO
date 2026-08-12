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

    [Fact]
    public void HtmlRender_EmitsEllipsisForOverflowingAtomicInlineContent() {
        const string html = "<div style='width:20px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis'><img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=' width='200' height='10'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText ellipsis = Assert.Single(EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>());

        Assert.Equal("\u2026", ellipsis.Text);
        Assert.InRange(ellipsis.Width, 0.01D, 20.01D);
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

        HtmlRenderText[] compactText = EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(compact).Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] wideText = EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(wide).Pages[0].Scene).OfType<HtmlRenderText>().ToArray();

        Assert.DoesNotContain('\t', string.Concat(compactText.Select(text => text.Text)));
        Assert.DoesNotContain('\t', string.Concat(wideText.Select(text => text.Text)));
        Assert.True(Assert.Single(wideText, text => text.Text == "B").X > Assert.Single(compactText, text => text.Text == "B").X);
        Assert.Contains("B", PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(wide).ToPdf(new HtmlPdfSaveOptions())).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRender_UsesInheritedAbsoluteLengthTabStops() {
        const string html = "<div style='font-size:10px;tab-size:32px'><pre style='margin:0;font-family:Consolas;font-size:20px;letter-spacing:.01px'>A\tB</pre></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlRenderText[] glyphs = EnumerateTextOverflowVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText a = Assert.Single(glyphs, glyph => glyph.Text == "A");
        HtmlRenderText b = Assert.Single(glyphs, glyph => glyph.Text == "B");

        Assert.InRange(b.X - a.X, 20D, 50D);
    }

    [Fact]
    public void HtmlRender_InheritedRelativeLengthTabStopsRetainTheParentsComputedLength() {
        const string html = "<div style='font-size:10px;tab-size:2em'><pre style='font-size:20px'>A\tB</pre></div>";
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed, new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions());
        HtmlRenderBoxStyle parent = resolver.Resolve(document.QuerySelector("div")!, 120D);
        HtmlRenderBoxStyle child = resolver.Resolve(document.QuerySelector("pre")!, 120D, parent);

        Assert.True(parent.TabSizeIsLength);
        Assert.True(child.TabSizeIsLength);
        Assert.Equal(20D, parent.TabSize, 3);
        Assert.Equal(20D, child.TabSize, 3);
    }

    [Fact]
    public void HtmlRender_TabSizeInitialResetsWhileInvalidDeclarationsRemainInherited() {
        const string html = "<div style='font-size:10px;tab-size:2em'><pre id='reset' style='tab-size:initial'>A\tB</pre><pre id='invalid' style='tab-size:-1px'>A\tB</pre><pre id='revert' style='tab-size:revert'>A\tB</pre></div>";
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed, new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions());
        HtmlRenderBoxStyle parent = resolver.Resolve(document.QuerySelector("div")!, 120D);
        HtmlRenderBoxStyle reset = resolver.Resolve(document.QuerySelector("#reset")!, 120D, parent);
        HtmlRenderBoxStyle invalid = resolver.Resolve(document.QuerySelector("#invalid")!, 120D, parent);
        HtmlRenderBoxStyle reverted = resolver.Resolve(document.QuerySelector("#revert")!, 120D, parent);

        Assert.False(reset.TabSizeIsLength);
        Assert.Equal(8D, reset.TabSize, 3);
        Assert.True(invalid.TabSizeIsLength);
        Assert.Equal(20D, invalid.TabSize, 3);
        Assert.True(reverted.TabSizeIsLength);
        Assert.Equal(20D, reverted.TabSize, 3);
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

    [Fact]
    public void HtmlRender_TextSpacingRetainsFiniteAuthoredLengthsOutsideFontRelativeRanges() {
        const string html = "<p style='margin:0;font-family:Consolas;font-size:16px;letter-spacing:200px;word-spacing:-20px'>A B</p>";
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed, new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        HtmlRenderBoxStyle style = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions()).Resolve(document.QuerySelector("p")!, 1000D);

        IReadOnlyList<HtmlRenderText> glyphs = EnumerateTextOverflowVisuals(HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 1000D,
            ViewportHeight = 100D
        }).Pages[0].Scene).OfType<HtmlRenderText>().ToList();

        Assert.Equal(200D, style.LetterSpacing, 3);
        Assert.Equal(-20D, style.WordSpacing, 3);
        Assert.Equal(new[] { "A", " ", "B" }, glyphs.Select(glyph => glyph.Text));
        Assert.True(glyphs[1].X - glyphs[0].X > 200D);
        Assert.True(glyphs[2].X - glyphs[1].X > 180D);
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
