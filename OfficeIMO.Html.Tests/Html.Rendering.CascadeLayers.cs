using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlCascadeLayers_ApplyNormalAndImportantPrecedenceBeforeSpecificity() {
        const string html = """
            <style>
              @layer reset, theme;
              @layer theme { #target { color:blue; background-color:blue !important; } }
              @layer reset { .target { color:red; background-color:red !important; } }
              .target { color:lime; background-color:lime !important; }
            </style>
            <div id="target" class="target">Layered</div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IElement target = document.QuerySelector("#target")!;
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[target];

        Assert.Equal("rgba(0, 255, 0, 1)", style.GetValue("color"));
        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("background-color"));
    }

    [Fact]
    public void HtmlCascadeLayers_ElementDeclarationOverridesInheritedValueFromOutsideLayers() {
        const string html = """
            <style>
              @layer components;
              body { color:red; font-size:12px; }
              @layer components { .title { color:blue; font-size:28px; } }
            </style>
            <h1 class="title">Layered heading</h1>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".title")!];

        Assert.Equal("rgba(0, 0, 255, 1)", style.GetValue("color"));
        Assert.Equal("28px", style.GetValue("font-size"));
    }

    [Fact]
    public void HtmlCascadeLayers_KeepNestedAndAnonymousLayerOrderDeterministic() {
        const string html = """
            <style>
              @layer framework {
                @layer base { #target { border-color:red; } }
                @layer components { .target { border-color:blue; } }
              }
              @layer { .target { outline-color:purple; } }
              @layer { #target { outline-color:orange; } }
            </style>
            <div id="target" class="target">Layered</div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(0, 0, 255, 1)", style.GetValue("border-color"));
        Assert.Equal("rgba(255, 165, 0, 1)", style.GetValue("outline-color"));
    }

    [Fact]
    public void HtmlCascadeLayers_KeepSublayersWithinTheirDeclaredParentOrder() {
        const string html = """
            <style>
              @layer framework, application;
              @layer framework {
                @layer reset, components;
                @layer components { #target { color:blue; } }
                #target { background-color:lime; }
              }
              @layer application { .target { color:red; background-color:red; } }
            </style>
            <div id="target" class="target">Layered</div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("background-color"));
    }

    [Fact]
    public void HtmlCascadeLayers_RevertLayerFallsBackPastDeclarationsInTheCurrentLayer() {
        const string html = """
            <style>
              @layer base, theme;
              @layer base { #target { color:red; } }
              @layer theme { #target { color:blue; } }
              @layer theme { #target { color:revert-layer; } }
            </style>
            <div id="target">Layered</div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCascadeLayers_ImportantRevertLayerRevealsTheNextImportantLayer() {
        const string html = """
            <style>
              @layer base, theme;
              @layer base { #target { color:red !important; } }
              @layer theme { #target { color:blue !important; } }
              @layer base { #target { color:revert-layer !important; } }
            </style>
            <div id="target">Layered</div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(0, 0, 255, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_CombinesParentListsAmpersandsAndImplicitDescendants() {
        const string html = """
            <style>
              .card, #unused, .panel {
                color:red;
                & > .title { color:blue; }
                .body { color:lime; }
              }
              .card > .title { color:orange; }
            </style>
            <section class="card"><strong class="title">Title</strong><span class="body">Body</span></section>
            <section class="panel"><strong class="title">Panel</strong></section>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.QuerySelector(".card")!].GetValue("color"));
        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector(".card > .title")!].GetValue("color"));
        Assert.Equal("rgba(0, 255, 0, 1)", styles[document.QuerySelector(".body")!].GetValue("color"));
        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector(".panel > .title")!].GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_CarriesParentSelectorsThroughNestedConditionalRules() {
        const string html = """
            <style>
              @layer base, enhancements;
              .card {
                @layer enhancements { & > .title { color:blue; } }
                @supports (display:grid) { .body { color:lime; } }
                @media screen { & > .media { color:purple; } }
              }
            </style>
            <section class="card"><strong class="title">Title</strong><span class="body">Body</span><span class="media">Media</span></section>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector(".title")!].GetValue("color"));
        Assert.Equal("rgba(0, 255, 0, 1)", styles[document.QuerySelector(".body")!].GetValue("color"));
        Assert.Equal("rgba(128, 0, 128, 1)", styles[document.QuerySelector(".media")!].GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesSourceOrderAroundConditionalBlocks() {
        const string html = """
            <style>
              .conditional-first { @media screen { color:red; } color:blue; }
              .conditional-last { color:blue; @media screen { color:red; } }
            </style>
            <span class="conditional-first">First</span><span class="conditional-last">Last</span>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector(".conditional-first")!].GetValue("color"));
        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.QuerySelector(".conditional-last")!].GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesLiteralAmpersandsInsideAttributeSelectors() {
        const string html = "<style>.card { &[data-code='A&B'] { color:red; } }</style><div class='card' data-code='A&B'>Matched</div>";
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".card")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCascadeLayersAndNesting_FlowThroughTheManagedSceneAndExporters() {
        const string html = """
            <style>
              @layer base, theme;
              @layer base { .card { background:red; } }
              @layer theme {
                .card {
                  background:blue;
                  & > .title { color:lime; }
                }
              }
            </style>
            <section id="card" class="card" style="width:60px;height:24px"><strong class="title">LayerMarker</strong></section>
            """;
        var options = new HtmlRenderOptions {
            ViewportWidth = 90D,
            ViewportHeight = 40D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), options);
        HtmlRenderShape fill = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "section#card" && item.Shape.FillColor.HasValue);
        HtmlRenderText text = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            item => item.Text == "LayerMarker");
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);

        Assert.Equal(OfficeColor.Blue, fill.Shape.FillColor);
        Assert.Equal(OfficeColor.Lime, text.Color);
        Assert.Contains("#0000ff", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#00ff00", svg, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(HtmlConversionDocument.Parse(html).ToPng(options));
    }
}
