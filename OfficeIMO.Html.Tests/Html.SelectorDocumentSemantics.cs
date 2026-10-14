using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(HtmlCssPropertyGrammarCollection.Name)]
public sealed class HtmlSelectorDocumentSemanticsTests {
    [Theory]
    [InlineData("", true)]
    [InlineData("<!doctype html>", false)]
    [InlineData("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">", false)]
    public void IdAndClassSelectorsRespectDocumentModeWithoutFoldingAttributeValues(string doctype, bool quirky) {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(doctype
            + "<div id='SAMPLEÅ' class='SAMPLEÅ'>Target</div>");
        HtmlElement target = document.QuerySelector("div")!;
        foreach (string selector in new[] { "#sampleÅ", ".sampleÅ", ":is(#sampleÅ, .absent)" }) {
            Assert.Equal(quirky, HtmlCssSelectorParser.Parse(selector).Selector!.Matches(target));
            Assert.Equal(quirky, target.Matches(selector));
            Assert.Equal(quirky, document.QuerySelector(selector) != null);
        }

        foreach (string selector in new[] { "#sampleå", ".sampleå", "[id='sampleÅ']", "[class='sampleÅ']" }) {
            Assert.False(HtmlCssSelectorParser.Parse(selector).Selector!.Matches(target));
            Assert.False(target.Matches(selector));
        }
        Assert.Same(target, document.QuerySelector("[id='SAMPLEÅ']"));
    }

    [Theory]
    [InlineData("", true)]
    [InlineData("<!doctype html>", false)]
    public void IndexedRulesAndPseudoRulesPaintUsingTheDocumentMode(string doctype, bool quirky) {
        string html = doctype + """
            <style>
              div { width:40px; height:30px; background:white; }
              #sample { background:red; }
              .sample { width:60px; }
              .sample::before { content:'Marker'; color:blue; }
              [class='sample'] { background:lime; }
            </style>
            <div id="SAMPLE" class="SAMPLE sample-other"></div>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlElement target = document.Document.QuerySelector("div")!;
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[target];
        Assert.Equal(quirky ? "60px" : "40px", style.GetValue("width"));
        Assert.Equal(quirky ? "rgba(255, 0, 0, 1)" : "rgba(255, 255, 255, 1)", style.GetValue("background-color"));

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 100,
            Margins = HtmlRenderMargins.All(0),
            AllowSystemFontFallback = false
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
        Assert.Equal(quirky ? OfficeColor.Red : OfficeColor.White, raster.GetPixel(20, 25));
        Assert.Equal(quirky, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Any(text => text.Text.Contains("Marker", StringComparison.Ordinal)));
    }

    [Fact]
    public void LanguageSelectorsHonorNamespacePrecedenceAndEmptyInheritanceStops() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("""
            <!doctype html><html lang="de"><body>
              <p id="html" lang="en" xml:lang="fr">HTML</p>
              <svg lang="en" xml:lang="fr">
                <text id="inherited">French</text>
                <g lang="it"><text id="svg-ordinary">Italian</text></g>
                <g xml:lang=""><text id="unknown">Unknown</text></g>
              </svg>
              <math lang="nl"><mi id="math-ordinary">German</mi></math>
              <math xml:lang="es"><mi id="math-xml">Spanish</mi></math>
            </body></html>
            """);
        AssertLanguage(document, "html", "en", "fr");
        AssertLanguage(document, "inherited", "fr", "en");
        AssertLanguage(document, "svg-ordinary", "it", "fr");
        AssertLanguage(document, "math-ordinary", "de", "nl");
        AssertLanguage(document, "math-xml", "es", "de");
        HtmlElement unknown = document.QuerySelector("#unknown")!;
        Assert.False(unknown.Matches(":lang(fr)"));
        Assert.False(unknown.Matches(":lang(de)"));
        Assert.False(HtmlCssSelectorParser.Parse(":lang(fr)").Selector!.Matches(unknown));
    }

    [Fact]
    public void LanguageSelectorsDriveStylesThroughTheNativeProjection() {
        const string html = """
            <!doctype html><style>
              svg:lang(fr) { color:red; }
              svg:lang(en) { color:blue; }
            </style><svg lang="en" xml:lang="fr"><text id="target">French</text></svg>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        var styles = HtmlComputedStyleEngine.Compute(document);
        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.Document.QuerySelector("#target")!].GetValue("color"));
    }

    [Theory]
    [InlineData("main > section.card:first-child > p:nth-child(2)")]
    [InlineData("p:not(.absent):is(.ready, [data-state='WAIT' i])")]
    [InlineData("section > p:first-child + p, section > p:last-child")]
    [InlineData("p[data-tokens~='beta']:lang(en)")]
    [InlineData("section:empty, p:last-of-type")]
    public void CompleteOwnedListsPreserveProviderResultsAndScope(string selector) {
        const string html = """
            <!doctype html><main><section class="card" lang="en-US">
              <p id="first">First</p><p id="second" class="ready" data-tokens="alpha beta">Second</p>
              <p id="third" data-state="wait">Third</p>
            </section><section id="empty"></section></main>
            """;
        HtmlDocument owned = HtmlDocumentEngine.Default.ParseDocument(html);
        var provider = new AngleSharp.Html.Parser.HtmlParser().ParseDocument(html);
        Assert.True(HtmlCssSelectorParser.ParseList(selector).IsSupported);
        Assert.Equal(provider.QuerySelectorAll(selector).Select(element => element.Id),
            owned.QuerySelectorAll(selector).Select(element => element.Id));
        Assert.Equal(provider.QuerySelector("section")!.QuerySelectorAll(selector).Select(element => element.Id),
            owned.QuerySelector("section")!.QuerySelectorAll(selector).Select(element => element.Id));
    }

    [Fact]
    public void UnsupportedSelectorsRetainProviderFallbackAndInvalidSelectorErrors() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<!doctype html><section><p class='card'></p><input type='checkbox' checked></section>");
        Assert.Equal(HtmlCssSelectorParseStatus.Unsupported, HtmlCssSelectorParser.ParseList(":checked").Status);
        Assert.Equal("input", document.QuerySelector(":checked")!.LocalName);
        Assert.Equal("p", document.QuerySelector("section")!.QuerySelector(":scope > p")!.LocalName);
        Assert.Equal("selector", Assert.Throws<ArgumentException>(() => document.QuerySelectorAll("[")).ParamName);
        Assert.Equal("selector", Assert.Throws<ArgumentException>(() => document.QuerySelector("p")!.Matches("[")).ParamName);
        Assert.Equal("selector", Assert.Throws<ArgumentException>(() => document.QuerySelectorAll("p, [")).ParamName);
    }

    [Fact]
    public void DetachedAndTemplateQueriesKeepTheirScopeAndSourceDocumentMode() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("""
            <section><p id="TARGET" class="CARD">Attached</p></section>
            <template><p id="TEMPLATE" class="CARD">Template</p></template>
            """).Clone();
        HtmlElement section = document.QuerySelector("section")!;
        HtmlElement target = section.QuerySelector("#target")!;
        section.Remove();
        Assert.Null(document.QuerySelector("#target"));
        Assert.Same(target, section.QuerySelector(".card"));
        Assert.True(target.Matches("section > #target"));
        Assert.True(target.Matches(":first-child"));
        Assert.Empty(target.QuerySelectorAll("#target"));

        HtmlElement template = document.QuerySelector("template")!;
        Assert.Null(document.QuerySelector("#template"));
        Assert.Empty(template.QuerySelectorAll(".card"));
        Assert.Equal("TEMPLATE", template.TemplateContent!.QuerySelector("#template")!.Id);
        Assert.True(template.TemplateContent.QuerySelector(".card")!.Matches(":first-child"));

        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(
            "<span id='FRAGMENT' class='CARD'></span>", section);
        Assert.Equal(HtmlDocumentMode.Quirks, fragment.Document.Mode);
        Assert.Equal("FRAGMENT", fragment.QuerySelector("#fragment.card")!.Id);
    }

    private static void AssertLanguage(HtmlDocument document, string id, string expected, string rejected) {
        HtmlElement target = document.QuerySelector("#" + id)!;
        Assert.True(HtmlCssSelectorParser.Parse(":lang(" + expected + ")").Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse(":lang(" + rejected + ")").Selector!.Matches(target));
        Assert.True(target.Matches(":lang(" + expected + ")"));
        Assert.False(target.Matches(":lang(" + rejected + ")"));
        Assert.Contains(target, document.QuerySelectorAll(":lang(" + expected + ")"));
    }
}
