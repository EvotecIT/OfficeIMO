using System.Text;
using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlSecurityLayoutBatchTests {
    [Theory]
    [InlineData("var(--outer, var(--inner, red))", true)]
    [InlineData("var(--outer, var(--inner, red)", false)]
    [InlineData("var(--outer, var(--inner, 'red'))", true)]
    public void SupportsVarSyntaxChecksNestedValues(string value, bool expected) {
        Assert.Equal(expected, HtmlCssCustomPropertyResolver.HasValidVarFunctionSyntax(value));
    }

    [Fact]
    public void SupportsVarSyntaxRejectsExcessiveNestingWithoutRepeatedScanning() {
        string value = string.Concat(Enumerable.Repeat("var(--a,", 1024)) + "red" +
            new string(')', 1024);

        Assert.False(HtmlCssCustomPropertyResolver.HasValidVarFunctionSyntax(value));
    }

    [Fact]
    public void CustomPropertyExpansionStopsBeforeExponentialOutput() {
        var values = new Dictionary<string, string>(StringComparer.Ordinal) { ["--p0"] = "x" };
        for (int index = 1; index <= 20; index++)
            values[$"--p{index}"] = $"var(--p{index - 1})var(--p{index - 1})";

        Assert.False(HtmlCssCustomPropertyResolver.TryResolve("var(--p20)",
            name => values.TryGetValue(name, out string? value) ? value : null, out _));
        Assert.True(HtmlCssCustomPropertyResolver.TryResolve("var(--p4)",
            name => values.TryGetValue(name, out string? value) ? value : null, out string resolved));
        Assert.Equal(new string('x', 16), resolved);
    }

    [Fact]
    public void TrustedStyleExpansionCanExceedTheUntrustedCharacterCeiling() {
        string family = new string('a', 262145);
        string html = "<style>p{--family:" + family + ";font-family:var(--family)}</style><p>x</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html, HtmlConversionDocumentOptions.CreateTrustedProfile());

        Assert.Equal(family, HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("p")!]
            .GetValue("font-family"));
        HtmlConversionDocument untrusted = HtmlConversionDocument.Parse(html);
        Assert.NotEqual(family, HtmlComputedStyleEngine.Compute(untrusted)[untrusted.Document.QuerySelector("p")!]
            .GetValue("font-family"));
    }

    [Fact]
    public void TrustedContainerStyleQueryResolvesLargeCustomProperties() {
        string large = new string('a', 262145);
        string html = "<style>@container style(--x:var(--y)){#target{color:red}}</style>" +
            "<section style='--x:" + large + ";--y:" + large + "'><p id='target'>x</p></section>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html, HtmlConversionDocumentOptions.CreateTrustedProfile());

        Assert.Equal("rgba(255, 0, 0, 1)", HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#target")!]
            .GetValue("color"));
    }

    [Fact]
    public void TrustedCustomPropertiesResolveFiniteDeepChains() {
        var values = new Dictionary<string, string>(StringComparer.Ordinal) { ["--p40"] = "blue" };
        for (int index = 39; index >= 0; index--) values[$"--p{index}"] = $"var(--p{index + 1})";
        string? Lookup(string name) => values.TryGetValue(name, out string? value) ? value : null;

        Assert.False(HtmlCssCustomPropertyResolver.TryResolve("var(--p0)", Lookup, out _));
        Assert.True(HtmlCssCustomPropertyResolver.TryResolve("var(--p0)", Lookup, out string resolved,
            enforceLimits: false));
        Assert.Equal("blue", resolved);
    }

    [Fact]
    public void ClipPathRejectsExcessivePolygonVerticesBeforeMaterializingThem() {
        string polygon = "polygon(" + string.Join(",", Enumerable.Repeat("1px 1px", 4097)) + ")";

        Assert.False(HtmlCssClipPathParser.IsSupportedSyntax(polygon));
    }

    [Fact]
    public void ClipPathCountsOnlyTopLevelPolygonCommas() {
        string polygon = "polygon(" + string.Join(",", Enumerable.Repeat("clamp(0px,1px,2px) 1px", 2049)) + ")";

        Assert.True(HtmlCssClipPathParser.IsSupportedSyntax(polygon));
    }

    [Fact]
    public void ClipPathRejectsOverflowingNegativeInsetWithoutThrowing() {
        Assert.False(HtmlCssClipPathParser.IsSupportedSyntax("inset(-1e308px)"));
    }

    [Fact]
    public void FootnotePlannerStopsWhenNoPageCanHoldTheNote() {
        const string html = "<style>@page{size:10px 10px;margin:0}" +
            "body,p,.note{margin:0;font-size:1px;line-height:1px}" +
            ".note{float:footnote}</style><p>x<span class='note'>y</span></p>";
        var options = new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            HonorCssPageRules = true,
            MaxPageCount = 3
        };

        Assert.Throws<InvalidOperationException>(() => HtmlRenderTestDriver.Render(html, options));
    }

    [Fact]
    public void SvgHrefNormalizationHandlesManyUnmatchedEndTags() {
        var html = new StringBuilder();
        for (int index = 0; index < 1024; index++) html.Append("<div>");
        for (int index = 0; index < 1024; index++) html.Append("</missing>");
        html.Append("<svg><image href='new' xlink:href='old'/></svg>");

        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html.ToString());

        var image = Assert.IsAssignableFrom<AngleSharp.Dom.IElement>(document.QuerySelector("image"));
        Assert.Equal("new", HtmlDocumentParser.GetExactAttributeValue(image, "href"));
        Assert.Equal("old", HtmlDocumentParser.GetExactAttributeValue(image, "xlink:href"));
    }

    [Fact]
    public void TextShadowParserCountsExcessLayersWithoutMaterializingThem() {
        string value = string.Join(",", Enumerable.Repeat("1px 2px red", 32));

        Assert.True(HtmlCssTextShadowParser.TryParse(value, 16, 16, 100, 100, 100, 100,
            OfficeColor.Black, 2, out IReadOnlyList<HtmlCssTextShadow> shadows, out int total));
        Assert.Equal(2, shadows.Count);
        Assert.Equal(32, total);
        Assert.False(HtmlCssTextShadowParser.TryParse(value + ",invalid", 16, 16, 100, 100,
            100, 100, OfficeColor.Black, 2, out _, out _));
        Assert.False(HtmlCssTextShadowParser.TryParse(string.Join(",", Enumerable.Repeat("1px 2px red", 65)),
            16, 16, 100, 100, 100, 100, OfficeColor.Black, 2, out _, out _));
    }

    [Fact]
    public void InheritedQuotesHaveAQuotedPairLimit() {
        string value = string.Concat(Enumerable.Repeat("'[' ']' ", 65));

        Assert.False(HtmlCssQuotes.TryParse(value, out _));
        Assert.True(HtmlCssQuotes.TryParse("'[' ']'", out HtmlCssQuotes quotes));
        Assert.Equal("[", quotes.OpeningAt(0));
    }

    [Fact]
    public void RegisteredPropertiesCannotMultiplyAcrossAllElementsWithoutACombinedBudget() {
        string registrations = string.Concat(Enumerable.Range(0, 20).Select(index =>
            $"@property --p{index} {{ syntax:'<number>'; inherits:false; initial-value:0; }}"));
        string html = "<style>" + registrations + "</style>" +
            string.Concat(Enumerable.Repeat("<span>x</span>", 6000));

        Assert.Equal("MaxCssDeclarations", Assert.Throws<HtmlDomLimitException>(
            () => HtmlComputedStyleEngine.Compute(html)).LimitSource);
    }

    [Fact]
    public void PropertyDescriptorsConsumeTheCssDeclarationBudget() {
        var options = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Untrusted };
        options.Limits.MaxCssDeclarations = 4;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@property --tone { syntax:'<color>'; inherits:false; initial-value:red;"
            + " unknown:1; unknown:2; unknown:3; }</style><p>x</p>", options);

        Assert.Equal("MaxCssDeclarations", Assert.Throws<HtmlDomLimitException>(
            () => HtmlComputedStyleEngine.Compute(document)).LimitSource);
    }

    [Fact]
    public void PropertyDescriptorsAtTheExactCssDeclarationBudgetRemainValid() {
        var options = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Untrusted };
        options.Limits.MaxCssDeclarations = 5;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>@property --tone { syntax:'<color>'; inherits:false; initial-value:red; unknown:1; unknown:2; }</style><p>x</p>", options);

        HtmlComputedStyleEngine.Compute(document);
    }

    [Fact]
    public void TrustedStylesCanRegisterMoreThanTheFormerFixedPropertyCap() {
        string registrations = string.Concat(Enumerable.Range(0, 257).Select(index =>
            $"@property --p{index} {{ syntax:'<color>'; inherits:false; initial-value:red; }}"));
        string html = "<style>" + registrations + "p{color:var(--p256)}</style><p>Text</p>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html, HtmlConversionDocumentOptions.CreateTrustedProfile());

        Assert.Equal("red", HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("p")!].GetValue("color"));
    }

    [Fact]
    public void TrustedRegistrationsCanUseLongSyntaxWithMoreThanSixteenAlternatives() {
        string syntax = string.Join("|", Enumerable.Repeat("<color>", 40));
        string html = "<style>@property --tone{syntax:'" + syntax +
            "';inherits:false;initial-value:red}div{--tone:blue}p{color:var(--tone)}</style>" +
            "<div><p>Text</p></div>";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            html, HtmlConversionDocumentOptions.CreateTrustedProfile());

        Assert.Equal("red", HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("p")!].GetValue("color"));
    }

    [Fact]
    public void ZeroWidthLeaderCannotExpandIntoAnUnboundedString() {
        string html = "<style>p::before{content:leader('" + new string('\u200b', 1000) +
            "')}</style><p>x</p>";

        Assert.Equal("MaxLeaderCharacters", Assert.Throws<HtmlDomLimitException>(
            () => HtmlRenderTestDriver.Render(html)).LimitSource);
    }

    [Fact]
    public void FirstLetterWhitespaceScanRespectsTheRenderOperationBudget() {
        string html = "<style>p::first-letter{color:red}</style><p>" +
            new string(' ', 5000) + "A</p>";

        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), Assert.Throws<HtmlDomLimitException>(
            () => HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
                MaxLayoutOperations = 1000
            })).LimitSource);
    }

    [Fact]
    public void FirstLineCanSplitALongCjkTokenWithinTheWorkBudget() {
        string html = "<style>p::first-line{color:red}</style><p style='width:32px'>" +
            new string('漢', 2048) + "</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            MaxLayoutOperations = 10000
        });
        Assert.Contains('漢', rendered.Text);
    }

    [Fact]
    public void FirstLineWithNegativeLetterSpacingFindsTheLaterFittingPrefix() {
        const string html = "<style>p::first-line{color:red;letter-spacing:-6px}</style>" +
            "<p style='width:20px;margin:0;font-family:Arial;font-size:12px;overflow-wrap:anywhere'>" +
            "WWWWWWiiiiiiiiWWWWWW</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        static IEnumerable<HtmlRenderVisual> Flatten(IEnumerable<HtmlRenderVisual> visuals) {
            foreach (HtmlRenderVisual visual in visuals) {
                yield return visual;
                IEnumerable<HtmlRenderVisual>? children = visual switch {
                    HtmlRenderClipGroup clip => clip.Visuals,
                    HtmlRenderPathClipGroup path => path.Visuals,
                    HtmlRenderEffectGroup effect => effect.Visuals,
                    HtmlRenderSemanticGroup semantic => semantic.Visuals,
                    HtmlRenderLogicalTextGroup logical => logical.Visuals,
                    _ => null
                };
                if (children == null) continue;
                foreach (HtmlRenderVisual child in Flatten(children)) yield return child;
            }
        }
        string firstLine = string.Concat(Flatten(rendered.Pages[0].Scene)
            .OfType<HtmlRenderText>()
            .Where(item => item.Color == OfficeColor.Red)
            .Select(item => item.Text));

        Assert.True(firstLine.Length >= 14, $"Expected the later fitting prefix; got '{firstLine}'.");
    }

    [Fact]
    public void FirstLineSearchKeepsSurrogatePairsWholeAtTheFormerPrefixBoundary() {
        string token = new string('\u200B', 16383) + "😀suffix";
        var callbackInputs = new List<string>();
        var options = new HtmlRenderOptions {
            TextHyphenationCallback = value => {
                callbackInputs.Add(value);
                return Array.Empty<int>();
            }
        };
        string html = "<style>p::first-line{color:red}</style>" +
            "<p style='width:1px;hyphens:auto;overflow-wrap:anywhere'>" + token + "</p>";

        HtmlRenderTestDriver.Render(html, options);

        Assert.Contains(token, callbackInputs);
        Assert.DoesNotContain(callbackInputs, value => value.Length > 0 && char.IsHighSurrogate(value[value.Length - 1]));
    }

    [Fact]
    public void InheritedOversizedQuotesProduceOneBoundedDiagnostic() {
        string html = "<style>body{quotes:'" + new string('x', 9000) +
            "' 'y'}p::before{content:open-quote}</style>" +
            string.Concat(Enumerable.Repeat("<p>x</p>", 100));

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        HtmlDiagnostic diagnostic = Assert.Single(rendered.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            && item.Detail.StartsWith("quotes=", StringComparison.Ordinal));
        Assert.True(diagnostic.Detail.Length <= 263);
    }

    [Fact]
    public void QuoteCacheCapacityDoesNotChangeValidQuoteRendering() {
        var html = new StringBuilder("<style>span::before{content:open-quote}");
        for (int index = 0; index < 1025; index++) {
            html.Append("#q").Append(index).Append("{quotes:'").Append(index).Append("' 'x'}");
        }
        html.Append("</style>");
        for (int index = 0; index < 1025; index++) {
            html.Append("<span id='q").Append(index).Append("'>z </span>");
        }

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html.ToString());
        Assert.Contains("1024", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, item =>
            item.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            && item.Detail.StartsWith("quotes=", StringComparison.Ordinal));
    }

    [Fact]
    public void ManySmallLeadersConsumeTheRenderWideBudget() {
        string html = "<style>p::before{content:leader('\u200b')}p{margin:0}</style>" +
            string.Concat(Enumerable.Repeat("<p>x</p>", 12));

        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutOperations), Assert.Throws<HtmlDomLimitException>(
            () => HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
                ViewportWidth = 100,
                Margins = HtmlRenderMargins.All(0),
                MaxLayoutOperations = 50000,
                MaxLeaderCharacters = 20000
            })).LimitSource);
    }

}
