using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Tests.Pdf;
using System.Text;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRendering_MarkerPseudoElementUsesListOrdinalStyleAndOutsideGeometryAcrossBackends() {
        const string html = "<style>ol{list-style-position:outside}li::marker{content:'[' counter(list-item,upper-roman) '] ';color:#ff0000;font-size:10px}</style>"
            + "<ol start='4'><li id='first'>First</li><li value='7'>SecondPdf</li></ol>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            ViewportHeight = 70D,
            Margins = HtmlRenderMargins.All(20D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] texts = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText[] markers = texts.Where(text => text.Source == "list-marker").ToArray();
        HtmlRenderText first = Assert.Single(texts, text => text.Text == "First");
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));

        Assert.Equal(new[] { "[IV] ", "[VII] " }, markers.Select(marker => marker.Text));
        Assert.All(markers, marker => Assert.Equal(OfficeColor.Red, marker.Color));
        Assert.True(markers[0].X < first.X);
        Assert.Contains("#FF0000", svg, StringComparison.Ordinal);
        Assert.Contains("SecondPdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(list-style-position:outside)"));
        Assert.True(HtmlComputedStyleEngine.TryParsePseudoElementSelector("li::marker", out string host, out HtmlPseudoElementKind kind));
        Assert.Equal("li", host);
        Assert.Equal(HtmlPseudoElementKind.Marker, kind);
    }

    [Fact]
    public void HtmlRendering_ListStylePositionSeparatesInsideAndOutsideMarkersAndContentNoneSuppressesMarker() {
        const string html = "<style>#suppressed::marker{content:none}</style>"
            + "<ol><li id='outside' style='list-style-position:outside'>Outside</li>"
            + "<li id='inside' style='list-style-position:inside'>Inside</li>"
            + "<li id='suppressed'>Suppressed</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 160D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(20D)
        });
        HtmlRenderText[] texts = EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>().ToArray();
        HtmlRenderText outsideBody = Assert.Single(texts, text => text.Text == "Outside");
        HtmlRenderText insideBody = Assert.Single(texts, text => text.Text == "Inside");
        HtmlRenderText[] markers = texts.Where(text => text.Source == "list-marker").ToArray();

        Assert.Equal(2, markers.Length);
        Assert.True(markers[0].X < outsideBody.X);
        Assert.True(markers[1].X < insideBody.X);
        Assert.DoesNotContain(rendered.Text.Split('\n'), text => text == "3. ");
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(list-style-position:hanging)"));
    }

    [Fact]
    public void HtmlRendering_ListStyleImageUsesSharedResourcePipelineAndFallsBackToTextMarker() {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(6, 4));
        string source = "data:image/png;base64," + imageData;
        string html = "<ul style=\"list-style-type:none;list-style-image:url('" + source + "');list-style-position:outside\"><li>ImagePdf</li></ul>"
            + "<ul style=\"list-style-image:url('data:image/png;base64,not-valid')\"><li>Fallback</li></ul>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 160D,
            ViewportHeight = 70D,
            Margins = HtmlRenderMargins.All(10D)
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderImage image = Assert.Single(visuals.OfType<HtmlRenderImage>());
        HtmlRenderText fallback = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Source == "list-marker");
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, options).Bytes);
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));

        Assert.Equal(6D, image.Width, 3);
        Assert.Equal(4D, image.Height, 3);
        Assert.Equal("• ", fallback.Text);
        Assert.Contains("data:image/png;base64", svg, StringComparison.Ordinal);
        Assert.Contains("ImagePdf", PdfCore.PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(list-style-image:url('marker.png'))"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(list-style-image:linear-gradient(red,blue))"));
    }

    [Theory]
    [InlineData("inside")]
    [InlineData("outside")]
    public void HtmlRendering_MarkerPseudoElementPreservesMixedTextAndImages(string position) {
        string imageData = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(5, 3));
        string html = "<style>li::marker{content:'[' url('data:image/png;base64," + imageData + "') ']';color:#123456}</style>"
            + "<ul style='list-style-position:" + position + "'><li id='mixed'>Mixed marker</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 180D,
            Margins = HtmlRenderMargins.All(10D)
        });
        IReadOnlyList<HtmlRenderVisual> visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToList();

        HtmlRenderImage image = Assert.Single(visuals.OfType<HtmlRenderImage>());
        Assert.Equal(5D, image.Width, 3);
        Assert.Equal(3D, image.Height, 3);
        Assert.Equal(new[] { "[", "]" }, visuals.OfType<HtmlRenderText>()
            .Where(text => text.Source == "li#mixed::marker")
            .Select(text => text.Text)
            .ToArray());
    }

    [Fact]
    public void HtmlRendering_MarkerPseudoElementParticipatesInLayerCascade() {
        const string html = "<style>@layer base,theme;@layer base{li::marker{content:'A ';color:red}}"
            + "@layer theme{li::marker{content:'B ';color:blue}}</style><ul><li>Layered</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        HtmlRenderText marker = Assert.Single(
            EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderText>(),
            text => text.Source == "list-marker");
        Assert.Equal("B ", marker.Text);
        Assert.Equal(OfficeColor.Blue, marker.Color);
    }

    [Fact]
    public void HtmlRendering_UsesCanonicalHtmlListOrdinals() {
        const string html = "<ol start='9x'><li>First</li><li value='12junk'>Second</li><li>Third</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { "9. ", "First", "12. ", "Second", "13. ", "Third" },
            rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("upper-roman", "IV. ")]
    [InlineData("lower-alpha", "d. ")]
    [InlineData("lower-greek", "δ. ")]
    [InlineData("decimal-leading-zero", "04. ")]
    [InlineData("cjk-decimal", "四、")]
    [InlineData("hiragana", "え、")]
    [InlineData("hiragana-iroha", "に、")]
    [InlineData("katakana", "エ、")]
    [InlineData("katakana-iroha", "ニ、")]
    public void HtmlRendering_FormatsStandardOrderedListCounterStyles(string style, string marker) {
        string html = "<ol start='4' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("cjk-decimal", 204, "二〇四、")]
    [InlineData("full-width", 204, "２０４. ")]
    [InlineData("cjk-heavenly-stem", 10, "癸、")]
    [InlineData("cjk-earthly-branch", 12, "亥、")]
    public void HtmlRendering_FormatsBoundedEastAsianCounterStyles(string style, int start, string marker) {
        string html = "<ol start='" + start + "' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("japanese-informal", 10, "十、")]
    [InlineData("japanese-formal", 101, "壱百壱、")]
    [InlineData("korean-hangul-formal", 6001, "육천일, ")]
    [InlineData("korean-hanja-informal", 101, "百一, ")]
    [InlineData("korean-hanja-formal", 11, "壹拾壹, ")]
    [InlineData("simp-chinese-informal", 101, "一百零一、")]
    [InlineData("simp-chinese-formal", 6001, "陆仟零壹、")]
    [InlineData("trad-chinese-informal", 10, "十、")]
    [InlineData("trad-chinese-formal", 99, "玖拾玖、")]
    [InlineData("cjk-ideographic", 6001, "六千零一、")]
    public void HtmlRendering_FormatsLonghandEastAsianCounterStyles(string style, int start, string marker) {
        string html = "<ol start='" + start + "' style='list-style-type:" + style + "'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("japanese-informal", -11, "マイナス十一")]
    [InlineData("korean-hangul-formal", -11, "마이너스 일십일")]
    [InlineData("simp-chinese-formal", -101, "负壹佰零壹")]
    [InlineData("trad-chinese-formal", -101, "負壹佰零壹")]
    [InlineData("japanese-formal", 10000, "一〇〇〇〇")]
    public void HtmlRendering_FormatsLonghandEastAsianRepresentations(string style, int value, string expected) {
        Assert.True(HtmlCounterStyleFormatter.TryFormat(value, style, out string formatted));
        Assert.Equal(expected, formatted);
    }

    [Theory]
    [InlineData("circle", "◦ ")]
    [InlineData("square", "▪ ")]
    [InlineData("'→'", "→ ")]
    public void HtmlRendering_FormatsUnorderedAndQuotedListMarkers(string style, string marker) {
        string html = "<ul style=\"list-style-type:" + style + "\"><li>Item</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("disc", "• ")]
    [InlineData("'→'", "→ ")]
    public void HtmlRendering_OrderedListsUseTheSuffixOfAnExplicitBulletLikeStyle(string style, string marker) {
        string html = "<ol style=\"list-style-type:" + style + "\"><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { marker, "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_IncrementsNumericMarkersOnUnorderedLists() {
        const string html = "<ul style='list-style-type:decimal'><li>First</li><li>Second</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "1. ", "First", "2. ", "Second" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_AdditiveCounterStyleFallsBackForNegativeAutomaticRange() {
        const string html = "<style>@counter-style tally{system:additive;additive-symbols:1 'I';fallback:decimal}</style><ol start='-1' style='list-style-type:tally'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "-1. ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_CounterFallbackUsesTheEffectiveStylesAffixes() {
        const string html = "<style>"
            + "@counter-style base{system:numeric;symbols:'0' '1' '2' '3' '4' '5' '6' '7' '8' '9';prefix:'<';suffix:'> '}"
            + "@counter-style builtin-fallback{system:fixed;symbols:'I';prefix:'[';suffix:'] ';fallback:decimal}"
            + "@counter-style custom-fallback{system:fixed;symbols:'I';prefix:'[';suffix:'] ';fallback:base}"
            + "@counter-style bullet-fallback{system:fixed;symbols:'I';fallback:disc}"
            + "</style><ol start='2' style='list-style-type:builtin-fallback'><li>Built in</li></ol>"
            + "<ol start='2' style='list-style-type:custom-fallback'><li>Custom</li></ol>"
            + "<ol start='2' style='list-style-type:bullet-fallback'><li>Bullet</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "2. ", "Built in", "<2> ", "Custom", "• ", "Bullet" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_FormatsAuthorDefinedCounterStyleMarkers() {
        const string html = """
            <style>
              @counter-style binary {
                system:numeric;
                symbols:"0" "1";
                pad:4 "0";
                prefix:"[";
                suffix:"] ";
              }
              ol { list-style-type:binary; }
            </style>
            <ol start="3"><li>Item</li></ol>
            """;

        var options = new HtmlRenderOptions();
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        HtmlCounterStyleRegistry registry = HtmlCounterStyleRegistry.Parse(document, options);
        Assert.True(registry.TryFormatMarker(3, "binary", out string directMarker));
        Assert.Equal("[0011] ", directMarker);

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);

        Assert.Equal(new[] { "[0011] ", "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("cyclic 2")]
    [InlineData("numeric 2")]
    [InlineData("alphabetic 2")]
    [InlineData("symbolic 2")]
    [InlineData("additive 2")]
    [InlineData("fixed nope")]
    [InlineData("fixed 2 3")]
    public void HtmlRendering_InvalidCounterStyleSystemArityUsesInitialSymbolicSystem(string system) {
        string html = "<style>@counter-style marks{system:" + system + ";symbols:'X' 'Y';additive-symbols:1 'I'}</style>"
            + "<ol start='3' style='list-style-type:marks'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "XX. ", "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("pad:bogus")]
    [InlineData("range:bogus")]
    [InlineData("negative:one two three")]
    [InlineData("prefix:one two")]
    [InlineData("suffix:one two")]
    [InlineData("additive-symbols:bogus")]
    [InlineData("system:bogus")]
    public void HtmlRendering_InvalidOptionalCounterStyleDescriptorsUseTheirInitialValues(string descriptor) {
        string html = "<style>@counter-style marks{system:cyclic;symbols:'X';" + descriptor + "}</style>"
            + "<ol style='list-style-type:marks'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "X. ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_InvalidLaterCounterDescriptorRetainsEarlierValidValue() {
        const string html = "<style>@counter-style marks{system:cyclic;symbols:'A';symbols:'';suffix:') ';suffix:bogus two}</style><ol style='list-style-type:marks'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "A) ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_CounterStylePaddingIncludesNegativeAffixes() {
        const string html = "<style>@counter-style signed{system:numeric;symbols:'0' '1' '2' '3' '4' '5' '6' '7' '8' '9';negative:'-';pad:3 '0';suffix:' '}</style><ol start='-1' style='list-style-type:signed'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "-01 ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_PreservesAuthorCounterStyleIdentifierCasing() {
        const string html = "<style>@counter-style MyStyle{system:cyclic;symbols:'X'}</style><ul style='list-style-type:MyStyle'><li>Item</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "X. ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_IgnoresCommentsBeforeCounterStyleDescriptors() {
        const string html = "<style>@counter-style marks{/* mode */system:cyclic;/* glyph */symbols:'X';/* ending */suffix:') '}</style><ul style='list-style-type:marks'><li>Item</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "X) ", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_AllowsEmptyCounterStyleAffixes() {
        const string html = "<style>@counter-style marks{system:cyclic;symbols:'X';prefix:'';suffix:'';negative:'' ''}</style><ul style='list-style-type:marks'><li>Item</li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { "X", "Item" }, rendered.Text.Split('\n'));
    }

    [Fact]
    public void HtmlRendering_AppliesCounterStyleConditionalRulesForTheActiveMedia() {
        const string html = """
            <style>
              @media screen { @counter-style screen-only { system:cyclic; symbols:"S"; } }
              @media print { @counter-style print-only { system:cyclic; symbols:"P"; } }
            </style>
            """;
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        var document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();

        HtmlCounterStyleRegistry registry = HtmlCounterStyleRegistry.Parse(document, options);

        Assert.True(registry.TryFormat(1, "print-only", out string printed));
        Assert.Equal("P", printed);
        Assert.False(registry.TryFormat(1, "screen-only", out _));
    }

    [Theory]
    [InlineData("@counter-style mark{system:cyclic;symbols:'U'}@layer themed{@counter-style mark{system:cyclic;symbols:'L'}}", "U. ")]
    [InlineData("@layer override,base;@layer base{@counter-style mark{system:cyclic;symbols:'B'}}@layer override{@counter-style mark{system:cyclic;symbols:'O'}}", "B. ")]
    [InlineData("@layer outer{@counter-style mark{system:cyclic;symbols:'D'}@layer child{@counter-style mark{system:cyclic;symbols:'C'}}}", "D. ")]
    public void HtmlRendering_CounterStylesHonorCascadeLayerPrecedence(string css, string expectedMarker) {
        string html = "<style>" + css + "</style><ol style='list-style-type:mark'><li>Item</li></ol>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(new[] { expectedMarker, "Item" }, rendered.Text.Split('\n'));
    }

    [Theory]
    [InlineData("<ol start='2147483647' style='list-style-type:symbols(symbolic &quot;x&quot;)'><li>Item</li></ol>")]
    [InlineData("<style>@counter-style huge{system:symbolic;symbols:'x'}</style><ol start='2147483647' style='list-style-type:huge'><li>Item</li></ol>")]
    [InlineData("<style>@counter-style huge{system:additive;additive-symbols:1 'x'}</style><ol start='2147483647' style='list-style-type:huge'><li>Item</li></ol>")]
    public void HtmlRendering_BoundsExpandedCounterRepresentationsAndUsesDecimalFallback(string html) {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions());

        Assert.Equal(new[] { "2147483647. ", "Item" }, rendered.Text.Split('\n'));
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.CounterRepresentationLimitExceeded);
    }

}
