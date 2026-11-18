using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    private static readonly HtmlRenderOptions CssLayoutOptions = new() {
        Mode = HtmlRenderMode.Continuous,
        ViewportWidth = 400D,
        Margins = HtmlRenderMargins.All(0D)
    };

    // 16px text: x-height 8px, ascent 12.8px, descent 3.2px under CSS 2.1 10.8.1.
    [Theory]
    [InlineData("baseline", 40D, -40D)]
    [InlineData("middle", 40D, -24D)]
    [InlineData("text-top", 40D, -12.8D)]
    [InlineData("text-bottom", 40D, -36.8D)]
    [InlineData("-5px", 10D, -5D)]
    [InlineData("super", 10D, -14.8D)]
    [InlineData("sub", 10D, -7.6D)]
    public void HtmlRenderer_PositionsInlineBlocksByVerticalAlign(string verticalAlign, double height, double topFromBaseline) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>Ab <span style='display:inline-block;width:20px;height:" + height.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "px;background:#ff0000;vertical-align:" + verticalAlign + "'></span> x</p>", CssLayoutOptions);

        Assert.Equal(topFromBaseline, VerticalAlignBox(rendered).Y - TextBaseline(rendered, "Ab"), 1);
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bottom")]
    public void HtmlRenderer_AlignsTallInlineBlocksToTheLineEdges(string verticalAlign) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>Ab <span style='display:inline-block;width:20px;height:40px;background:#ff0000;vertical-align:"
            + verticalAlign + "'></span> x</p>", CssLayoutOptions);
        double baseline = TextBaseline(rendered, "Ab");

        // The 40px box sets the 40px line and spans it; the text sits at the opposite edge.
        Assert.Equal(0D, VerticalAlignBox(rendered).Y, 1);
        Assert.True(verticalAlign == "top" ? baseline < 20D : baseline > 30D, "baseline " + baseline);
    }

    [Fact]
    public void HtmlRenderer_ShiftsInlineTextByVerticalAlignKeywords() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>Ab <span style='font-size:8px;vertical-align:middle'>mid</span>"
            + " <span style='font-size:32px'>BIG</span> <span style='font-size:8px;vertical-align:top'>top</span>"
            + " <span style='font-size:8px;vertical-align:bottom'>bot</span></p>", CssLayoutOptions);
        double baseline = TextBaseline(rendered, "Ab");

        // middle: 0.3 x 8px above the 8px box centre, against half the parent's 8px x-height.
        Assert.Equal(baseline - 1.6D, TextBaseline(rendered, "mid"), 1);
        Assert.True(TextBaseline(rendered, "top") < baseline - 10D);
        Assert.True(TextBaseline(rendered, "bot") > baseline);
    }

    [Theory]
    [InlineData("<p style='text-align:justify;width:200px;margin:0'>漢字の文章は単語の間に空白がないので文字の間で両端揃えにする必要があります。</p>", 200D)]
    [InlineData("<p style='text-align:justify;text-justify:inter-character;width:300px;margin:0'>" + JustifiedText + "</p>", 300D)]
    public void HtmlRenderer_JustifiesBetweenCharacters(string html, double width) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssLayoutOptions);
        HtmlRenderText[] firstLine = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text))
            .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First()
            .OrderBy(text => text.X).ToArray();

        // Characters are painted one by one, the line ends at the right edge, and no approximation is reported.
        Assert.True(firstLine.Length > 8);
        Assert.Equal(width, firstLine.Max(text => text.X + (text.TextAdvanceWidth ?? text.Width)), 1);
        Assert.DoesNotContain(rendered.NativeDiagnostics, item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported);
    }

    [Fact]
    public void HtmlRenderer_InterWordJustificationKeepsWordsWhole() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;text-justify:inter-word;width:200px;margin:0'>漢字の文章は単語の間に空白がないので文字の間で両端揃えにする必要があります。 end</p>",
            CssLayoutOptions);

        Assert.DoesNotContain(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text == "字");
    }

    [Theory]
    [InlineData("full-width", "ab 12", "ａｂ　１２")]
    [InlineData("uppercase full-width", "ab", "ＡＢ")]
    [InlineData("full-width", "ｶﾞ", "ガ")]
    [InlineData("full-size-kana", "ァィゥッ ｧ", "アイウツ ｱ")]
    public void HtmlRenderer_AppliesWidthAndKanaTextTransforms(string transform, string source, string expected) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;text-transform:" + transform + "'>" + source + "</p>", CssLayoutOptions);

        Assert.Equal(expected, string.Concat(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>().Select(text => text.Text)));
        Assert.DoesNotContain(rendered.NativeDiagnostics, item => item.Code == HtmlRenderDiagnosticCodes.CssPropertyUnsupported);
    }

    [Theory]
    [InlineData("petite-caps")]
    [InlineData("all-petite-caps")]
    public void HtmlRenderer_SynthesizesPetiteCapsLikeSmallCaps(string variant) {
        HtmlRenderText[] smallCaps = CapsPieces("small-caps".Replace("small", variant.StartsWith("all", StringComparison.Ordinal) ? "all-small" : "small"));
        HtmlRenderText[] petiteCaps = CapsPieces(variant);

        Assert.Equal(smallCaps.Select(text => (text.Text, text.Font.Size)), petiteCaps.Select(text => (text.Text, text.Font.Size)));
        // Without petite-cap glyphs the letters are synthesized at the small-cap size.
        Assert.Contains(petiteCaps, text => text.Font.Size < 16D);

        static HtmlRenderText[] CapsPieces(string caps) =>
            JustificationVisuals(HtmlRenderTestDriver.Render("<p style='margin:0;font-variant-caps:" + caps + "'>Petite</p>", CssLayoutOptions)
                .Pages[0].Visuals).OfType<HtmlRenderText>().ToArray();
    }

    [Theory]
    [InlineData("letter-spacing:50%", "letter-spacing:10px")]
    [InlineData("word-spacing:50%", "word-spacing:10px")]
    public void HtmlRenderer_ResolvesPercentageSpacingAgainstTheFontSize(string percentage, string length) {
        double[] Positions(string declaration) =>
            JustificationVisuals(HtmlRenderTestDriver.Render("<p style='margin:0;font-size:20px;" + declaration + "'>ab cd</p>", CssLayoutOptions)
                .Pages[0].Visuals).OfType<HtmlRenderText>().Select(text => Math.Round(text.X, 3)).ToArray();

        Assert.Equal(Positions(length), Positions(percentage));
    }

    [Theory]
    [InlineData("white-space-collapse:preserve-breaks", "one   two\nthree", new[] { "one two", "three" })]
    [InlineData("white-space-collapse:preserve-spaces", "a  b\nc", new[] { "a  b c" })]
    [InlineData("text-wrap:nowrap;width:80px", "aaa bbb ccc ddd", new[] { "aaa bbb ccc ddd" })]
    [InlineData("white-space:normal;text-wrap-mode:nowrap;width:80px", "aaa bbb ccc ddd", new[] { "aaa bbb ccc ddd" })]
    // A later white-space resets both of its longhands.
    [InlineData("text-wrap-mode:nowrap;white-space:normal;width:80px", "aaa bbb ccc ddd", new[] { "aaa bbb", "ccc ddd" })]
    public void HtmlRenderer_AppliesWhiteSpaceAndTextWrapLonghands(string declarations, string text, string[] expected) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='" + declarations + "'>" + text + "</div>", CssLayoutOptions);

        Assert.Equal(expected, JustifiedLines(rendered).Select(line => line.Text).ToArray());
    }

    [Fact]
    public void HtmlRenderer_PreformattedTextWrapsWithTextWrapMode() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<pre style='text-wrap-mode:wrap;width:80px;margin:0'>aaa bbb ccc ddd</pre>", CssLayoutOptions);

        Assert.True(JustifiedLines(rendered).Count > 1);
    }

    [Fact]
    public void HtmlRenderer_BalancesShortBlocksWithoutAddingLines() {
        const string heading = "A heading that is long enough to wrap onto a second short line";
        IReadOnlyList<JustifiedLine> Lines(string wrap) => JustifiedLines(HtmlRenderTestDriver.Render(
            "<h2 style='width:300px;margin:0;font-size:20px;text-align:center;" + wrap + "'>" + heading + "</h2>", CssLayoutOptions));
        IReadOnlyList<JustifiedLine> ordinary = Lines(string.Empty);
        IReadOnlyList<JustifiedLine> balanced = Lines("text-wrap:balance");

        Assert.Equal(ordinary.Count, balanced.Count);
        double Spread(IReadOnlyList<JustifiedLine> lines) => lines.Max(line => line.Right - line.Left) - lines.Min(line => line.Right - line.Left);
        Assert.True(Spread(balanced) < Spread(ordinary) / 2D, "Balanced lines must have similar lengths.");
        // Alignment still uses the full width.
        Assert.All(balanced, line => Assert.Equal(300D - line.Right, line.Left, 1));
    }

    [Theory]
    [InlineData("<div align='center' style='width:300px'><div style='width:50px;background:#ff0000'>b</div></div>", 125D)]
    [InlineData("<div align='right' style='width:300px'><div style='width:50px;background:#ff0000'>b</div></div>", 250D)]
    [InlineData("<div style='width:300px;text-align:-webkit-center'><div style='width:50px;background:#ff0000'>b</div></div>", 125D)]
    // Nested blocks inherit the legacy alignment; authored margins still win.
    [InlineData("<div align='center' style='width:300px'><div><div style='width:50px;background:#ff0000'>b</div></div></div>", 125D)]
    [InlineData("<div align='center' style='width:300px'><div style='width:50px;margin-left:10px;background:#ff0000'>b</div></div>", 10D)]
    public void HtmlRenderer_LegacyDivAlignmentAlignsChildBlocks(string html, double expectedX) {
        HtmlRenderShape box = Assert.Single(JustificationVisuals(HtmlRenderTestDriver.Render(html, CssLayoutOptions).Pages[0].Visuals)
            .OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 });

        Assert.Equal(expectedX, box.X, 1);
    }

    private static HtmlRenderShape VerticalAlignBox(HtmlRenderDocument rendered) =>
        Assert.Single(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 });

    // Positioned text paints its baseline one font size below its frame, plus any baseline shift.
    private static double TextBaseline(HtmlRenderDocument rendered, string text) {
        HtmlRenderText visual = Assert.Single(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(),
            item => item.Text.Trim() == text);
        return visual.Y + visual.BaselineOffset + visual.Font.Size;
    }
}
