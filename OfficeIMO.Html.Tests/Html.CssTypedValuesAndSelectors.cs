using System.Text.Json;
using System.Threading;
using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(HtmlCssPropertyGrammarCollection.Name)]
public sealed class HtmlCssTypedValuesAndSelectorsTests {
    [Fact]
    public void IndependentTypedValueCorpusMatchesTheDeclaredSlice() {
        TypedValueCase[] corpus = Read<TypedValueCase>("css-typed-values-corpus.json");
        Assert.Equal(22, corpus.Length);
        foreach (TypedValueCase item in corpus) {
            HtmlCssPropertyParseResult result = HtmlCssPropertyParser.Parse(item.Property, item.Value);
            Assert.Equal(ParseEnum<HtmlCssPropertyParseStatus>(item.Status), result.Status);
            if (item.Kind != null) Assert.Equal(ParseEnum<HtmlCssPropertyValueKind>(item.Kind), result.Value!.Kind);
            if (item.Canonical != null) Assert.Equal(item.Canonical, result.Value!.CanonicalText);
            if (item.NumericType != null) {
                Assert.Equal(ParseEnum<HtmlCssNumericType>(item.NumericType), result.Value!.NumericValue!.Type);
                Assert.Equal(item.Number!.Value, result.Value.NumericValue.Value, 10);
            }
            if (item.ColorFunction != null) {
                Assert.Equal(ParseEnum<HtmlCssColorFunctionKind>(item.ColorFunction), result.Value!.ColorFunction!.Kind);
                Assert.Equal(item.Legacy, result.Value.ColorFunction.UsesLegacyCommaSyntax);
                Assert.Equal(3, result.Value.ColorFunction.Components.Count);
            }
        }
    }

    [Fact]
    public void IndependentSelectorCorpusMatchesOwnedDomAndClassifiesFallback() {
        SelectorCase[] corpus = Read<SelectorCase>("css-selector-corpus.json");
        Assert.Equal(16, corpus.Length);
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("""
            <body><main><header></header><aside></aside><article id="target" class="card primary" data-role="hero" data-tags="alpha beta" lang="en-US"></article></main></body>
            """);
        OfficeIMO.Html.Dom.HtmlElement target = document.QuerySelector("#target")!;
        foreach (SelectorCase item in corpus) {
            HtmlCssSelectorParseResult result = HtmlCssSelectorParser.Parse(item.Selector);
            Assert.Equal(ParseEnum<HtmlCssSelectorParseStatus>(item.Status), result.Status);
            if (item.Specificity != null) Assert.Equal(item.Specificity, result.Selector!.Specificity.ToString());
            if (item.Match != null) Assert.True(result.Selector!.Matches(target), item.Name + " should match the target.");
        }
    }

    [Fact]
    public void OwnedSelectorsDriveTheManagedCascadeAndUnsupportedSelectorsFallBack() {
        const string html = """
            <style>
              main > article.card[data-role='hero' i] { color:hsl(120 100% 25% / 50%); opacity:calc(.2 + .3); }
              main > article.card[data-fallback='MATCH' i] { background:red; visibility:collapse; opacity:.25; }
              article:first-child { visibility:hidden !important; }
            </style>
            <main><article id="target" class="card" data-role="HERO" data-fallback="match">Target</article></main>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document,
            new HtmlComputedStyleOptions { IncludeCascadeTraces = true })[document.Document.QuerySelector("#target")!];

        Assert.Equal("rgba(0, 128, 0, 0.5)", style.GetValue("color"));
        Assert.Equal("0.25", style.GetValue("opacity"));
        Assert.Equal("hidden", style.GetValue("visibility"));
        HtmlCssCascadeTrace opacityTrace = style.GetCascadeTrace("opacity")!;
        Assert.All(opacityTrace.Candidates, candidate => Assert.Equal(HtmlCssPropertyParseStatus.Parsed, candidate.GrammarStatus));
        Assert.Contains(opacityTrace.Candidates, candidate => candidate.Decision == HtmlCssCascadeDecision.Selected
            && candidate.DeclaredValue == "0.25");
    }

    [Fact]
    public void ModernNumericHslAndHwbReachDeterministicComputedSrgbValues() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <style>
              #hsl { color:hsl(120 .5 .5 / 25%); }
              #hwb { color:hwb(120 .2 .3); }
            </style>
            <p id="hsl"></p><p id="hwb"></p>
            """);
        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles =
            HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(1, 1, 1, 0.25)", styles[document.Document.QuerySelector("#hsl")!].GetValue("color"));
        Assert.Equal("rgba(1, 254, 1, 1)", styles[document.Document.QuerySelector("#hwb")!].GetValue("color"));
    }

    [Fact]
    public void SelectorParsingHonorsCancellationAndStructuralBounds() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlCssSelectorParser.Parse("article.card", cancellationToken: cancellation.Token));
        Assert.Equal(nameof(HtmlCssSelectorOptions.MaxCompounds), Assert.Throws<HtmlCssSelectorLimitException>(() =>
            HtmlCssSelectorParser.Parse("a b", new HtmlCssSelectorOptions { MaxCompounds = 1 })).LimitName);
        Assert.Equal(nameof(HtmlCssSelectorOptions.MaxSimpleSelectors), Assert.Throws<HtmlCssSelectorLimitException>(() =>
            HtmlCssSelectorParser.Parse("a.x", new HtmlCssSelectorOptions { MaxSimpleSelectors = 1 })).LimitName);
    }

    [Fact]
    public void OwnedSelectorMatchingPreservesCommentsMultipleIdsAndHtmlAttributeCasing() {
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<form id='same' class='panel' method='POST' data-empty='value'></form>");
        OfficeIMO.Html.Dom.HtmlElement target = document.QuerySelector("form")!;

        Assert.True(HtmlCssSelectorParser.Parse("form/**/.panel#same#same[method=post]").Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("#same#other").Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("[data-empty^='']").Selector!.Matches(target));
        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue,
            HtmlCssPropertyParser.Parse("opacity", "(.5)").Status);
        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue,
            HtmlCssPropertyParser.Parse("opacity", ".2 + .3").Status);
    }

    [Fact]
    public void OwnedSelectorMatchingUsesAsciiCaseAndEmptyNamespaceAttributes() {
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<main><article id='target' data-value='GRÜN'></article></main>").Edit(editor =>
                editor.QuerySelector("#target")!.SetAttribute(new OfficeIMO.Html.Dom.HtmlAttribute(
                    "qualified:att", "value", "urn:example")));
        OfficeIMO.Html.Dom.HtmlElement target = document.QuerySelector("#target")!;

        Assert.False(HtmlCssSelectorParser.Parse("[data-value='grün' i]").Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("[att]").Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("[qualified\\:att]").Selector!.Matches(target));
    }

    [Fact]
    public void OwnedMathAndSelectorTraversalStayBoundedOnAdversarialInputs() {
        string invalidUnary = "calc(" + string.Concat(Enumerable.Repeat("+ ", 10_000)) + "1)";
        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue,
            HtmlCssPropertyParser.Parse("opacity", invalidUnary).Status);

        const int depth = 80;
        string html = string.Concat(Enumerable.Repeat("<div>", depth)) + "target"
            + string.Concat(Enumerable.Repeat("</div>", depth));
        OfficeIMO.Html.Dom.HtmlElement deepest = HtmlDocumentEngine.Default.ParseDocument(html)
            .QuerySelectorAll("div").Last();
        HtmlCssSelector selector = HtmlCssSelectorParser.Parse(
            ".missing " + string.Join(" ", Enumerable.Repeat("*", depth - 1))).Selector!;
        Assert.False(selector.Matches(deepest));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => selector.Matches(deepest, cancellation.Token));
    }

    [Fact]
    public void OwnedSelectorTraversalConsumesTheConversionEvaluationBudget() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 4;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>.missing * *{color:red}</style><main><section><p>target</p></section></main>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlComputedStyleEngine.Compute(document));
        Assert.Equal(nameof(HtmlConversionLimits.MaxSelectorEvaluations), exception.LimitSource);
    }

    private static T[] Read<T>(string name) => JsonSerializer.Deserialize<T[]>(File.ReadAllText(
        Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", name)),
        new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;

    private static T ParseEnum<T>(string value) where T : struct =>
        (T)Enum.Parse(typeof(T), value);

    private sealed class TypedValueCase {
        public string Name { get; set; } = string.Empty;
        public string Property { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string? Kind { get; set; }
        public string? Canonical { get; set; }
        public string? NumericType { get; set; }
        public double? Number { get; set; }
        public string? ColorFunction { get; set; }
        public bool Legacy { get; set; }
    }

    private sealed class SelectorCase {
        public string Name { get; set; } = string.Empty;
        public string Selector { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string? Specificity { get; set; }
        public string? Match { get; set; }
    }
}
