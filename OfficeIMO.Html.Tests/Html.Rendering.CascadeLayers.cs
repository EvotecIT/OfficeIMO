using AngleSharp.Dom;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlCascadeLayers_InlineImportantOutranksLayeredAuthorImportant() {
        const string html = "<style>@layer base { #target { color:blue !important; } }</style><p id='target' style='color:red !important'>Inline</p>";

        HtmlRenderText text = Assert.Single(HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderText>(), item => item.Text == "Inline");

        Assert.Equal(OfficeColor.Red, text.Color);
    }

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
    public void HtmlCascadeLayers_UnlayeredRevertLayerRollsBackTheAuthorOrigin() {
        const string html = "<style>@layer base { p { color:blue; margin-left:12px; } } p { color:revert-layer; margin-left:revert-layer; }</style><p>Rollback</p>";
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("p")!];

        Assert.Equal(string.Empty, style.GetValue("color"));
        Assert.Equal(string.Empty, style.GetValue("margin-left"));
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
    public void HtmlCssNesting_UsesOwnedNamespacesAndTypedDeclarations() {
        const string html = """
            <style>
              @namespace svg url("http://www.w3.org/2000/svg");
              svg|svg.card {
                & > svg|a:first-child { color:hwb(120 0% 0%); width:calc(10px + 5px); }
              }
            </style>
            <svg class="card"><a id="target">Owned</a></svg>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(0, 255, 0, 1)", style.GetValue("color"));
        Assert.Equal("calc(10px + 5px)", style.GetValue("width"));
        Assert.True(style.TryGetTypedValue("width", out OfficeIMO.Html.Css.HtmlCssPropertyValue? width));
        Assert.Equal(OfficeIMO.Html.Css.HtmlCssPropertyValueKind.Calculation, width!.Kind);
    }

    [Fact]
    public void HtmlCssNesting_DoesNotPairInactiveConditionalRulesWithLaterSelectors() {
        const string html = """
            <style>
              @media print { .card { & > .title { color:red; } } }
              .card { & > .title { color:blue; } }
            </style>
            <section class="card"><strong class="title">Screen</strong></section>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".title")!];

        Assert.Equal("rgba(0, 0, 255, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesDeclarationsInterleavedWithQualifiedRules() {
        const string html = """
            <style>
              .last-declaration { color:green; & { color:blue; } color:red; }
              .last-rule { color:green; & { color:blue; } }
            </style>
            <p class="last-declaration">Red</p><p class="last-rule">Blue</p>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.QuerySelector(".last-declaration")!].GetValue("color"));
        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector(".last-rule")!].GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_RetainsProviderFallbackForUnsupportedPseudoClasses() {
        const string html = """
            <style>.form { & input:required { color:red; } }</style>
            <form class="form"><input id="target" required></form>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesDeclarationOrderWhenNestedSelectorUsesProviderFallback() {
        const string html = """
            <style>.field { color:green; &:where(:required) { color:blue; } color:red; }</style>
            <input class="field" required>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".field")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesDeclarationOrderAcrossUnknownNestedAtRules() {
        const string html = """
            <style>.field { color:green; @future ignored { color:purple; } & { color:blue; } color:red; }</style>
            <input class="field">
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".field")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_KeepsOwnedNamespaceEnvelopeAroundProviderPseudoClass() {
        const string html = """
            <style>
              @namespace h url("http://www.w3.org/1999/xhtml");
              h|form.form { & > h|input:required { color:red; } }
            </style>
            <form class="form"><input id="target" required></form>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector("#target")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCssNesting_DoesNotSplitEscapedOrCommentedCommas() {
        const string html = """
            <style>
              .host { & > .item\,special { color:red; } }
              .host { & > .commented/*,*/.target { background:blue; } }
            </style>
            <div class="host"><span id="escaped" class="item,special">Escaped</span><span id="commented" class="commented target">Commented</span></div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.QuerySelector("#escaped")!].GetValue("color"));
        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.QuerySelector("#commented")!].GetValue("background-color"));
    }

    [Fact]
    public void HtmlCssNesting_PreservesFunctionalProviderPseudoSpecificity() {
        const string html = """
            <style>
              .host { &:has(#marker) { color:red; } }
              .host.alt { color:blue; }
            </style>
            <div class="host alt"><span id="marker">Marker</span></div>
            """;
        var document = HtmlDocumentParser.ParseDocument(html);

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.QuerySelector(".host")!];

        Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
    }

    [Fact]
    public void HtmlCascadeLayers_TreatCommentsAsWhitespaceAroundRevertLayer() {
        const string normalHtml = """
            <style>
              @layer base, theme;
              @layer base { #target { color:red; } }
              @layer theme { #target { color:blue; } }
              @layer theme { #target { color:/**/revert-layer/**/; } }
            </style>
            <p id="target">Normal</p>
            """;
        const string importantHtml = """
            <style>
              @layer base, theme;
              @layer base { #target { color:red !important; } }
              @layer theme { #target { color:blue !important; } }
              @layer base { #target { color:/**/revert-layer/**/!important; } }
            </style>
            <p id="target">Important</p>
            """;
        var normalDocument = HtmlDocumentParser.ParseDocument(normalHtml);
        var importantDocument = HtmlDocumentParser.ParseDocument(importantHtml);

        Assert.Equal("rgba(255, 0, 0, 1)", HtmlComputedStyleEngine.Compute(normalDocument)[normalDocument.QuerySelector("#target")!].GetValue("color"));
        Assert.Equal("rgba(0, 0, 255, 1)", HtmlComputedStyleEngine.Compute(importantDocument)[importantDocument.QuerySelector("#target")!].GetValue("color"));
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
