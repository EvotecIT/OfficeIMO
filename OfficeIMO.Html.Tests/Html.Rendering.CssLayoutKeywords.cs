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

    // The box aligns against its 40px parent span: x-height 20px, ascent 32px, descent 8px.
    [Theory]
    [InlineData("middle", -15D)]
    [InlineData("text-top", -32D)]
    [InlineData("text-bottom", -2D)]
    public void HtmlRenderer_AlignsNestedInlineBlocksAgainstTheirParentFont(string verticalAlign, double topFromBaseline) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>Ab <span style='font-size:40px'>Xy<span style='display:inline-block;width:20px;height:10px;"
            + "background:#ff0000;vertical-align:" + verticalAlign + "'></span></span></p>", CssLayoutOptions);

        Assert.Equal(topFromBaseline, VerticalAlignBox(rendered).Y - TextBaseline(rendered, "Ab"), 1);
    }

    [Fact]
    public void HtmlRenderer_AlignsAgainstAZeroSizedParentFont() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>Ab <span style='font-size:0'><span style='display:inline-block;width:20px;height:10px;"
            + "font-size:16px;background:#ff0000;vertical-align:middle'></span></span></p>", CssLayoutOptions);

        // middle centres the box on the baseline plus half the parent's x-height, which is zero here.
        Assert.Equal(-5D, VerticalAlignBox(rendered).Y - TextBaseline(rendered, "Ab"), 1);
    }

    [Fact]
    public void HtmlRenderer_KeepsBaselinesInsideTopAlignedInlines() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px'>x <span style='vertical-align:top;font-size:32px'>A<span style='font-size:8px'>B</span></span>"
            + " <span style='vertical-align:bottom;font-size:32px'>C<span style='font-size:8px'>D</span></span></p>", CssLayoutOptions);

        Assert.Equal(TextBaseline(rendered, "A"), TextBaseline(rendered, "B"), 1);
        Assert.Equal(TextBaseline(rendered, "C"), TextBaseline(rendered, "D"), 1);
    }

    // An atomic box keeps its baseline relationship with the text of the top or bottom inline it sits in.
    [Theory]
    [InlineData("top", "display:inline-block;width:10px;height:10px")]
    [InlineData("bottom", "display:inline-block;width:10px;height:10px")]
    [InlineData("top", "display:inline-block;width:10px;height:10px;vertical-align:middle")]
    public void HtmlRenderer_AtomicBoxesJoinTheirInlineAlignmentGroup(string edge, string box) {
        string Html(string boxStyle) => "<p style='margin:0;font-size:16px'><span style='font-size:64px'>X</span>"
            + "<span style='vertical-align:" + edge + "'>A<span style='" + boxStyle + ";background:#ff0000'></span></span></p>";
        HtmlRenderDocument grouped = HtmlRenderTestDriver.Render(Html(box), CssLayoutOptions);
        HtmlRenderDocument alone = HtmlRenderTestDriver.Render("<p style='margin:0;font-size:16px'>A<span style='" + box
            + ";background:#ff0000'></span></p>", CssLayoutOptions);

        // The box sits where it would against A alone, however far the group moved from the paragraph baseline.
        Assert.Equal(VerticalAlignBox(alone).Y - TextBaseline(alone, "A"), VerticalAlignBox(grouped).Y - TextBaseline(grouped, "A"), 1);
    }

    [Fact]
    public void HtmlRenderer_PercentageVerticalAlignRaisesAtomicBoxesByTheirLineHeight() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px;line-height:20px'>Ab <span style='display:inline-block;width:10px;height:10px;"
            + "background:#ff0000;vertical-align:50%'></span></p>", CssLayoutOptions);

        // 50% of the 20px line-height lifts the box's bottom edge 10px above the baseline.
        Assert.Equal(-20D, VerticalAlignBox(rendered).Y - TextBaseline(rendered, "Ab"), 1);
    }

    [Fact]
    public void HtmlRenderer_ShiftedInlineBoxesBesideFloatsPushTheNextLineDown() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<div style='width:300px'><div style='float:left;width:50px;height:300px'></div>"
            + "<p style='margin:0;font-size:16px;line-height:20px'>Ab <span style='display:inline-block;width:10px;height:10px;"
            + "vertical-align:-100px;background:#ff0000'></span> one two three four five six seven eight nine ten eleven twelve</p></div>",
            CssLayoutOptions);
        HtmlRenderShape box = VerticalAlignBox(rendered);
        HtmlRenderText first = Assert.Single(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text.Trim() == "Ab");
        HtmlRenderText[] laterLines = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text) && text.LayoutY > first.LayoutY + 1D).ToArray();

        // The box hangs 100px below the first baseline, so the first line is that deep and the next starts below it.
        Assert.NotEmpty(laterLines);
        Assert.All(laterLines, text => Assert.True(text.LayoutY >= box.Y + box.Height - 0.5D, text.Text + " at " + text.LayoutY));
    }

    [Fact]
    public void HtmlRenderer_AlignsBeforeAndAfterAsSeparateGroups() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::before{content:'A';vertical-align:top;line-height:10px}p::after{content:'Z';vertical-align:bottom;line-height:40px}</style>"
            + "<p style='margin:0;font-size:16px;line-height:20px'>x</p>", CssLayoutOptions);

        // The 40px ::after sets the line and sits at its bottom; the 10px ::before stays at the top.
        Assert.True(TextBaseline(rendered, "Z") - TextBaseline(rendered, "A") > 10D,
            "A " + TextBaseline(rendered, "A") + ", Z " + TextBaseline(rendered, "Z"));
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

    [Theory]
    [InlineData("")]
    // Right-to-left text elsewhere in the paragraph sends every segment through bidi resolution.
    [InlineData(" \u05E9\u05DC\u05D5\u05DD")]
    public void HtmlRenderer_JustifiesLetterSpacedTextBetweenEveryCharacter(string suffix) {
        double[] FirstWord(string align) {
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
                "<p style='" + align + "letter-spacing:1px;width:300px;margin:0'>" + JustifiedText + suffix + "</p>", CssLayoutOptions);
            return JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
                .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First()
                .OrderBy(text => text.X).Take(5).Select(text => text.X).ToArray();
        }
        double[] start = FirstWord(string.Empty);
        double[] justified = FirstWord("text-align:justify;text-justify:inter-character;");

        // Each gap inside "Lorem" receives the same share, so the n-th letter moves n shares.
        double share = justified[1] - start[1];
        Assert.True(share > 0.01D, "share " + share);
        for (int index = 2; index < 5; index++) Assert.Equal(index * share, justified[index] - start[index], 2);
    }

    [Fact]
    public void HtmlRenderer_JustifiedLinesKeepTheSpaceOfBidiControls() {
        // The left-to-right mark owns the gap before the following letter, and the Hebrew word sends the line through bidi.
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;text-justify:inter-character;width:300px;margin:0'>Lo\u200E" + JustifiedText.Substring(2)
            + " \u05E9\u05DC\u05D5\u05DD</p>", CssLayoutOptions);
        HtmlRenderText[] firstLine = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text))
            .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First().ToArray();

        Assert.Equal(300D, firstLine.Max(text => text.X + (text.TextAdvanceWidth ?? text.Width)), 1);
    }

    // A and V kern in this font, so painting each letter by its own width would overshoot the measured line.
    [Theory]
    [InlineData("text-align-last:justify;text-justify:inter-character")]
    [InlineData("text-align:right;letter-spacing:2px")]
    public void HtmlRenderer_PerCharacterPaintKeepsKerning(string declarations) {
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        };
        options.Fonts.Add(OfficeIMO.TestAssets.ManagedTextShapingTestAssets.FamilyName,
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -120, includeSpace: true));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style=\"font-family:'" + OfficeIMO.TestAssets.ManagedTextShapingTestAssets.FamilyName + "';font-size:20px;width:300px;margin:0;"
            + declarations + "\">AVAVAV AVAV</p>", options);
        HtmlRenderText[] texts = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text)).OrderBy(text => text.X).ToArray();

        Assert.Contains(texts, text => text.Text == "V");
        Assert.Equal(300D, texts.Max(text => text.X + (text.TextAdvanceWidth ?? text.Width)), 1);
    }

    [Fact]
    public void HtmlRenderer_InterWordJustificationKeepsWordsWhole() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align:justify;text-justify:inter-word;width:200px;margin:0'>漢字の文章は単語の間に空白がないので文字の間で両端揃えにする必要があります。 end</p>",
            CssLayoutOptions);

        Assert.DoesNotContain(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(), text => text.Text == "字");
    }

    [Theory]
    // Spaces keep collapsing and line breaking; preserved white space that does not wrap becomes ideographic.
    [InlineData("full-width", " ab   12 ", "ａｂ １２")]
    [InlineData("full-width;white-space:pre", "ab  12", "ａｂ　　１２")]
    [InlineData("uppercase full-width", "ab", "ＡＢ")]
    [InlineData("full-width", "ｶﾞ", "ガ")]
    // Each run's own white space decides whether its spaces become ideographic.
    [InlineData("full-width", "<span style='white-space:pre'>a b</span><span> c   d </span>", "ａ　ｂ ｃ ｄ")]
    [InlineData("full-width", "<span> c   d </span><span style='white-space:pre'>a b</span>", "ｃ ｄ ａ　ｂ")]
    // Halfwidth forms map to their direct wide counterparts, here the compatibility jamo.
    [InlineData("full-width", "\uFFA1\uFFC2", "\u3131\u314F")]
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
    // Children inherit the resolved longhands, not the declaration that lost.
    [InlineData("text-wrap-mode:nowrap;white-space:normal;width:80px", "<span>aaa bbb ccc ddd</span>", new[] { "aaa bbb", "ccc ddd" })]
    [InlineData("white-space-collapse:preserve;white-space:normal", "<span>a   b</span>", new[] { "a b" })]
    // An explicit inherit of a shorthand takes the parent's longhands and outranks earlier ones.
    [InlineData("white-space-collapse:preserve", "<span style='white-space-collapse:collapse;white-space:inherit'>a   b</span>", new[] { "a   b" })]
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
    // Flex items keep their own alignment.
    [InlineData("<div align='center' style='width:300px;display:flex'><div style='width:50px;background:#ff0000'>b</div><div style='width:50px'>c</div></div>", 0D)]
    public void HtmlRenderer_LegacyDivAlignmentAlignsChildBlocks(string html, double expectedX) {
        HtmlRenderShape box = Assert.Single(JustificationVisuals(HtmlRenderTestDriver.Render(html, CssLayoutOptions).Pages[0].Visuals)
            .OfType<HtmlRenderShape>(), shape => shape.Shape.FillColor is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 });

        Assert.Equal(expectedX, box.X, 1);
    }

    [Fact]
    public void HtmlRenderer_JustifiesGapsBetweenElementsByTheirCommonAncestor() {
        const string words = " lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor</p>";
        HtmlRenderText[] FirstLine(string align) =>
            JustificationVisuals(HtmlRenderTestDriver.Render("<p style='" + align + "width:200px;margin:0;font-size:16px'>"
                    + "<span style='text-justify:inter-character'>aa</span><span>aa</span>" + words, CssLayoutOptions)
                .Pages[0].Visuals).OfType<HtmlRenderText>()
                .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First()
                .OrderBy(text => text.X).ToArray();
        HtmlRenderText[] start = FirstLine(string.Empty);
        HtmlRenderText[] justified = FirstLine("text-align:justify;text-justify:inter-word;");
        double glyph = (start[1].X - start[0].X) / 2D;

        // The first span widens the gap inside it; the paragraph's inter-word keeps the gap after it.
        Assert.Equal(new[] { "a", "a", "aa" }, justified.Take(3).Select(text => text.Text));
        Assert.True(justified[1].X - justified[0].X > glyph + 0.01D, "inner gap " + (justified[1].X - justified[0].X));
        Assert.Equal(glyph, justified[2].X - justified[1].X, 2);
    }

    [Fact]
    public void HtmlRenderer_JustifiesGapsByAnAncestorThatOwnsNoText() {
        HtmlRenderText[] firstLine = JustificationVisuals(HtmlRenderTestDriver.Render(
                "<p style='text-align:justify;text-justify:inter-word;width:200px;margin:0;font-size:16px'>"
                + "<span style='text-justify:inter-character'><span>aa</span><span>aa</span></span>"
                + " lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor</p>", CssLayoutOptions)
            .Pages[0].Visuals).OfType<HtmlRenderText>()
            .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First()
            .OrderBy(text => text.X).Take(4).ToArray();

        // The wrapper's inter-character governs the gap between its two children as well as the gaps inside them.
        Assert.Equal(new[] { "a", "a", "a", "a" }, firstLine.Select(text => text.Text));
        Assert.Equal(firstLine[1].X - firstLine[0].X, firstLine[2].X - firstLine[1].X, 2);
        Assert.Equal(firstLine[1].X - firstLine[0].X, firstLine[3].X - firstLine[2].X, 2);
    }

    [Fact]
    public void HtmlRenderer_FirstLineKeepsTheTextJustifyOfItsText() {
        HtmlRenderText[] firstLine = JustificationVisuals(HtmlRenderTestDriver.Render(
                "<style>p::first-line{color:#0000ff}</style><p style='text-align:justify;text-justify:inter-word;width:200px;margin:0;font-size:16px'>"
                + "<span style='text-justify:inter-character'>aaaa</span>"
                + " lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor</p>", CssLayoutOptions)
            .Pages[0].Visuals).OfType<HtmlRenderText>()
            .GroupBy(text => Math.Round(text.LayoutY, 1)).OrderBy(group => group.Key).First()
            .OrderBy(text => text.X).ToArray();

        // ::first-line cannot set text-justify, so the span's inter-character still spreads its letters.
        Assert.Equal(new[] { "a", "a", "a", "a" }, firstLine.Take(4).Select(text => text.Text));
    }

    [Fact]
    public void HtmlRenderer_SeparatorsFollowTheirOwnTextUnderAnAncestorNone() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align-last:justify;text-justify:none;width:200px;margin:0;font-size:16px'>"
            + "<span style='text-justify:inter-word'>A </span><span style='text-justify:inter-word'>B</span></p>", CssLayoutOptions);

        // The space belongs to the first span's inter-word text, so B is pushed to the right edge.
        Assert.Equal(200D, Assert.Single(JustifiedLines(rendered)).Right, 1);
    }

    [Fact]
    public void HtmlRenderer_DistributeJustifiesLikeInterCharacter() {
        double[] Positions(string mode) => JustificationVisuals(HtmlRenderTestDriver.Render(
                "<p style='text-align-last:justify;text-justify:" + mode + ";width:200px;margin:0'>abcdef</p>", CssLayoutOptions)
            .Pages[0].Visuals).OfType<HtmlRenderText>().Select(text => Math.Round(text.X, 2)).ToArray();

        Assert.Equal(Positions("inter-character"), Positions("distribute"));
    }

    [Fact]
    public void HtmlRenderer_InterCharacterKeepsCursiveWordsWhole() {
        const string arabic = "\u0645\u0631\u062D\u0628\u0627";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='text-align-last:justify;text-justify:inter-character;width:300px;margin:0'>" + arabic + " " + arabic + "</p>",
            CssLayoutOptions);

        // Arabic letters join, so the space goes between the words and each word is painted whole.
        Assert.DoesNotContain(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(),
            text => text.Text.Length == 1 && arabic.Contains(text.Text));
        Assert.Equal(300D, Assert.Single(JustifiedLines(rendered)).Right, 1);
    }

    [Fact]
    public void HtmlRenderer_NoJustificationAtGapsWhoseAncestorDisablesIt() {
        HtmlRenderText[] texts = JustificationVisuals(HtmlRenderTestDriver.Render(
                "<p style='text-align:justify;text-align-last:justify;text-justify:none;width:200px;margin:0;font-size:16px'>"
                + "<span style='text-justify:inter-character'>\u6F22</span><span style='text-justify:inter-character'>\u5B57</span></p>",
                CssLayoutOptions).Pages[0].Visuals).OfType<HtmlRenderText>().OrderBy(text => text.X).ToArray();

        // Neither span has a gap of its own, and the paragraph's none governs the one between them.
        Assert.True(texts[1].X < 40D, "second character at " + texts[1].X);
    }

    [Fact]
    public void HtmlRenderer_BalanceKeepsTheLineCountWhenFirstLineStylingChangesWidths() {
        string words = string.Join(" ", Enumerable.Repeat("aa", 12));
        IReadOnlyList<JustifiedLine> Lines(string wrap) => JustifiedLines(HtmlRenderTestDriver.Render(
            "<style>div::first-line{font-size:20px}</style><div style='width:300px;font-size:40px;" + wrap + "'>" + words + "</div>",
            CssLayoutOptions));

        Assert.Equal(Lines(string.Empty).Count, Lines("text-wrap:balance").Count);
    }

    [Fact]
    public void HtmlRenderer_BalanceLeavesLongBlocksUnchanged() {
        string text = JustifiedText + " " + JustifiedText + " " + JustifiedText;
        IReadOnlyList<JustifiedLine> Lines(string wrap) => JustifiedLines(HtmlRenderTestDriver.Render(
            "<p style='width:200px;margin:0;" + wrap + "'>" + text + "</p>", CssLayoutOptions));
        IReadOnlyList<JustifiedLine> ordinary = Lines(string.Empty);

        // Balancing is limited to short blocks, as browsers limit it.
        Assert.True(ordinary.Count > 6, "lines " + ordinary.Count);
        Assert.Equal(ordinary.Select(line => line.Text), Lines("text-wrap:balance").Select(line => line.Text));
    }

    [Fact]
    public void HtmlRenderer_BalancedBlocksMatchFirstLineAtTheBalancedLength() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>h2::first-line{color:#ff0000}</style><h2 style='width:300px;margin:0;font-size:20px;text-wrap:balance'>"
            + "A heading that is long enough to wrap onto a second short line</h2>", CssLayoutOptions);
        HtmlRenderText[] texts = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(text => !string.IsNullOrWhiteSpace(text.Text)).ToArray();
        double firstLineY = texts.Min(text => Math.Round(text.LayoutY, 1));

        Assert.True(texts.Select(text => Math.Round(text.LayoutY, 1)).Distinct().Count() > 1);
        Assert.All(texts, text => Assert.Equal(Math.Round(text.LayoutY, 1) == firstLineY,
            text.Color is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 }));
    }

    [Fact]
    public void HtmlRenderer_JustifiesCharactersOfTextAProviderWouldShape() {
        // keep-all keeps each CJK word in one segment, which a configured provider would shape as one glyph run.
        const string text = "漢字の文章 単語の間に 空白がない 文字の間で 両端揃えに する必要が あります";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test';text-align:justify;text-justify:inter-character;word-break:keep-all;width:205px;margin:0\">"
            + text + "</p>");
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D),
            TextShapingProvider = new OfficeIMO.TestAssets.ManagedTextShapingTestAssets.RecordingProvider()
        };
        options.Fonts.Add(OfficeIMO.TestAssets.ManagedTextShapingTestAssets.FamilyName,
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFont(text.Distinct().Select(character => (int)character).ToArray()));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(document, options);
        HtmlRenderText[] firstLine = JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>()
            .Where(item => !string.IsNullOrWhiteSpace(item.Text))
            .GroupBy(item => Math.Round(item.LayoutY, 1)).OrderBy(group => group.Key).First()
            .OrderBy(item => item.X).ToArray();

        // Each character is placed with its share up to the right edge, and the dropped provider glyph run is reported.
        Assert.All(firstLine, item => Assert.Equal(1, item.Text.Length));
        Assert.Equal(205D, firstLine.Max(item => item.X + (item.TextAdvanceWidth ?? item.Width)), 1);
        Assert.Contains(rendered.NativeDiagnostics, item => item.Code == HtmlRenderDiagnosticCodes.ComplexTextShapingUnsupported
            && item.Detail == "justified-provider-glyph-context-not-retained");
    }

    [Fact]
    public void HtmlRenderer_EmptyInlineLineHeightRaisesLinesWithTopAlignedBoxes() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<p style='margin:0;font-size:16px;line-height:20px'>Ab <span style='line-height:60px'></span>"
            + "<span style='display:inline-block;width:10px;height:10px;vertical-align:top'></span></p>"
            + "<div style='height:5px;background:#ff0000'></div>", CssLayoutOptions);

        Assert.True(VerticalAlignBox(rendered).Y >= 59.9D, "line bottom " + VerticalAlignBox(rendered).Y);
    }

    [Fact]
    public void HtmlRenderer_BalancedBlocksResolvePercentagesAndLeadersAgainstTheirWidth() {
        HtmlRenderText[] Visuals(string html) => JustificationVisuals(HtmlRenderTestDriver.Render(html, CssLayoutOptions)
            .Pages[0].Visuals).OfType<HtmlRenderText>().ToArray();

        HtmlRenderText first = Visuals("<p style='width:200px;text-wrap:balance;text-indent:20%;margin:0'>A<br>B</p>")
            .Single(text => text.Text == "A");
        HtmlRenderText label = Visuals("<style>p::after{content:leader('.') 'Z'}</style>"
            + "<p style='width:200px;text-wrap:balance;margin:0'>A<br>B</p>").Single(text => text.Text == "Z");
        HtmlRenderText[] clipped = Visuals("<p style='width:200px;text-wrap:balance;overflow:hidden;text-overflow:ellipsis;"
            + "margin:0;font-size:16px'>A<br>WWWWWWWW</p>");

        Assert.Equal(40D, first.X, 1);
        Assert.Equal(200D, label.X + (label.TextAdvanceWidth ?? label.Width), 1);
        Assert.DoesNotContain(clipped, text => text.Text.Contains('…'));
    }

    [Fact]
    public void HtmlRenderer_PreserveSpacesTurnsTabsIntoSpaces() {
        double BoxWidth(string text) => Assert.Single(JustificationVisuals(HtmlRenderTestDriver.Render(
                "<p style='margin:0;white-space-collapse:preserve-spaces;tab-size:8'><span style='display:inline-block;background:#ff0000'>"
                + text + "</span></p>", CssLayoutOptions).Pages[0].Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Shape.FillColor is OfficeIMO.OfficeColor { R: 0xFF, G: 0, B: 0 }).Width;

        Assert.Equal(BoxWidth("A B"), BoxWidth("A\tB"), 2);
    }

    // Each value gives 4px on a 40px child of a 20px parent; em parts keep the parent's font.
    [Theory]
    [InlineData("letter-spacing", "10%")]
    [InlineData("word-spacing", "10%")]
    [InlineData("letter-spacing", "calc(10%)")]
    [InlineData("word-spacing", "max(10%, 1px)")]
    [InlineData("letter-spacing", "calc(0.1em + 5%)")]
    public void HtmlRenderer_InheritedPercentageSpacingFollowsTheChildFont(string property, string value) {
        double[] Positions(string parent, string child) =>
            JustificationVisuals(HtmlRenderTestDriver.Render("<p style='margin:0;font-size:20px;" + parent + "'><span style='font-size:40px;"
                    + child + "'>ab cd</span></p>", CssLayoutOptions)
                .Pages[0].Visuals).OfType<HtmlRenderText>().Select(text => Math.Round(text.X, 3)).ToArray();

        Assert.Equal(Positions(string.Empty, property + ":4px"), Positions(property + ":" + value, string.Empty));
    }

    [Fact]
    public void HtmlRenderer_GeneratedSegmentBreaksBecomeSpacesBeforeTextTransform() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<style>p::before{content:'A\\A B'}</style>"
            + "<p style='margin:0;white-space-collapse:preserve-spaces;text-wrap-mode:nowrap;text-transform:full-width'></p>", CssLayoutOptions);

        Assert.Equal("Ａ　Ｂ", string.Concat(JustificationVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>().Select(text => text.Text)));
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
