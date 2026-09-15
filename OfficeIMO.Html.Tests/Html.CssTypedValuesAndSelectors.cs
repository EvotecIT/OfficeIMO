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
        Assert.Equal(25, corpus.Length);
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
               article:lang(en) { white-space:pre-wrap; }
            </style>
            <main><article id="target" class="card" lang="en" data-role="HERO" data-fallback="match">Target</article></main>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document,
            new HtmlComputedStyleOptions { IncludeCascadeTraces = true })[document.Document.QuerySelector("#target")!];

        Assert.Equal("rgba(0, 128, 0, 0.5)", style.GetValue("color"));
        Assert.Equal("0.25", style.GetValue("opacity"));
        Assert.Equal("hidden", style.GetValue("visibility"));
        Assert.Equal("pre-wrap", style.GetValue("white-space"));
        HtmlCssCascadeTrace opacityTrace = style.GetCascadeTrace("opacity")!;
        Assert.All(opacityTrace.Candidates, candidate => Assert.Equal(HtmlCssPropertyParseStatus.Parsed, candidate.GrammarStatus));
        Assert.Contains(opacityTrace.Candidates, candidate => candidate.Decision == HtmlCssCascadeDecision.Selected
            && candidate.DeclaredValue == "0.25");
    }

    [Fact]
    public void OwnedSelectorListsStructuralPseudosAndLogicalPseudosMatchWithCssSpecificity() {
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("""
            <html><body><main>
              <section id="first"><p class="only">one</p></section>
              <section id="middle"><p>one</p><em>two</em><p id="target" class="picked">three</p><p>four</p></section>
              <section id="last"><!-- comment --></section>
            </main></body></html>
            """);
        OfficeIMO.Html.Dom.HtmlElement target = document.QuerySelector("#target")!;

        HtmlCssSelectorListParseResult list = HtmlCssSelectorParser.ParseList(
            "#missing, main > section:nth-child(2) > p:nth-of-type(2):is(.picked, strong):not(.excluded)");
        Assert.True(list.IsSupported);
        Assert.Equal(2, list.SelectorList!.Selectors.Count);
        Assert.True(list.SelectorList.Matches(target));
        Assert.Equal("0,4,3", list.SelectorList.Selectors[1].Specificity.ToString());
        Assert.True(HtmlCssSelectorParser.Parse("section:first-child p:only-child").Selector!.Matches(document.QuerySelector(".only")!));
        Assert.True(HtmlCssSelectorParser.Parse("section:last-child:empty").Selector!.Matches(document.QuerySelector("#last")!));
        Assert.True(HtmlCssSelectorParser.Parse("p:nth-last-child(2)").Selector!.Matches(target));
        Assert.True(HtmlCssSelectorParser.Parse("p:nth-last-of-type(even)").Selector!.Matches(target));
        Assert.Equal("0,0,0", HtmlCssSelectorParser.Parse(":where(#target, .picked)").Selector!.Specificity.ToString());
    }

    [Fact]
    public void OwnedNamespaceSelectorsUseStylesheetBindingsAndExpandedNames() {
        const string svgNamespace = "http://www.w3.org/2000/svg";
        const string xlinkNamespace = "http://www.w3.org/1999/xlink";
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<body><svg><a id='target'></a></svg><a id='html'></a></body>").Edit(editor => {
                editor.QuerySelector("#target")!.SetAttribute(new OfficeIMO.Html.Dom.HtmlAttribute("xlink:href", "asset.svg", xlinkNamespace));
                editor.Body!.AppendChild(editor.CreateElement("a", string.Empty));
            });
        OfficeIMO.Html.Dom.HtmlElement target = document.QuerySelector("#target")!;
        HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(
            "@namespace svg url('http://www.w3.org/2000/svg');@namespace xlink 'http://www.w3.org/1999/xlink';svg|a[xlink|href$='.svg']{}");
        HtmlCssNamespaceContext namespaces = HtmlCssNamespaceContext.FromStyleSheet(sheet);
        var options = new HtmlCssSelectorOptions { Namespaces = namespaces };

        Assert.True(HtmlCssSelectorParser.Parse("svg|a[xlink|href]", options).Selector!.Matches(target));
        Assert.True(HtmlCssSelectorParser.Parse("*|a[*|href]", options).Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("|a", options).Selector!.Matches(target));
        Assert.False(HtmlCssSelectorParser.Parse("|a", options).Selector!.Matches(document.QuerySelector("#html")!));
        Assert.True(HtmlCssSelectorParser.Parse("|a", options).Selector!.Matches(document.Body!.Children.Last()));
        Assert.Equal(svgNamespace, namespaces.Prefixes["svg"]);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse("missing|a", options).Status);
        Assert.True(HtmlCssSelectorParser.Parse("svg/**/|/**/a[xlink/**/|/**/href]", options).Selector!.Matches(target));
        Assert.Null(HtmlCssNamespaceContext.FromStyleSheet(
            HtmlCssSyntaxParser.ParseStyleSheet("a{}@namespace 'urn:late';")).DefaultNamespaceUri);
    }

    [Fact]
    public void LogicalPseudoArgumentsDoNotInheritTheStylesheetDefaultNamespace() {
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument(
            "<body><div class='ancestor'><p id='html' class='picked'></p></div>"
            + "<svg><g class='ancestor'><a id='svg' class='picked'></a></g><a id='plain'></a></svg></body>");
        var options = new HtmlCssSelectorOptions {
            Namespaces = new HtmlCssNamespaceContext("http://www.w3.org/2000/svg")
        };
        OfficeIMO.Html.Dom.HtmlElement html = document.QuerySelector("#html")!;
        OfficeIMO.Html.Dom.HtmlElement svg = document.QuerySelector("#svg")!;
        OfficeIMO.Html.Dom.HtmlElement plain = document.QuerySelector("#plain")!;

        Assert.True(HtmlCssSelectorParser.Parse("*|*:is(.picked)", options).Selector!.Matches(html));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:where(.picked)", options).Selector!.Matches(html));
        Assert.False(HtmlCssSelectorParser.Parse("*|*:not(.picked)", options).Selector!.Matches(html));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:is(.picked)", options).Selector!.Matches(svg));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:not(.picked)", options).Selector!.Matches(plain));
        Assert.False(HtmlCssSelectorParser.Parse("*|*:is(.ancestor .picked)", options).Selector!.Matches(html));
        Assert.False(HtmlCssSelectorParser.Parse("*|*:where(.ancestor .picked)", options).Selector!.Matches(html));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:not(.ancestor .picked)", options).Selector!.Matches(html));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:is(.ancestor .picked)", options).Selector!.Matches(svg));
        Assert.True(HtmlCssSelectorParser.Parse("*|*:where(.ancestor .picked)", options).Selector!.Matches(svg));
        Assert.False(HtmlCssSelectorParser.Parse("*|*:not(.ancestor .picked)", options).Selector!.Matches(svg));
    }

    [Fact]
    public void GroupedPseudoAndNamespaceSelectorsDriveTheManagedCascade() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <style>
              @namespace svg url("http://www.w3.org/2000/svg");
              .picked, main > p:last-child { color: blue; }
              main > p:nth-child(even):is(.picked, .other):not(.excluded) { color: red; }
              :where(#target) { color: green; }
              svg|a:last-child { visibility: hidden; string-set: marker 'Vector'; }
              p:lang(en) { white-space: pre-wrap; }
            </style>
            <main><p>first</p><p id="target" class="picked" lang="en">second</p></main>
            <svg><a id="svg-target" lang="en">link</a></svg>
            """);
        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.Document.QuerySelector("#target")!].GetValue("color"));
        Assert.Equal("pre-wrap", styles[document.Document.QuerySelector("#target")!].GetValue("white-space"));
        Assert.Equal("hidden", styles[document.Document.QuerySelector("#svg-target")!].GetValue("visibility"));
        Assert.Equal("marker \"Vector\"", styles[document.Document.QuerySelector("#svg-target")!].GetValue("string-set"));
    }

    [Fact]
    public void NamespaceQualifiedSelectorsRetainProviderFallbackAndListSpecificity() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <style>
              @namespace svg url("http://www.w3.org/2000/svg");
              svg|a:lang(en), .fallback { color: blue; }
              svg|a:lang(en) { string-set: provider 'Vector'; }
              #vector { color: red; }
              p.fallback { background-color: red; }
              #missing, :where(.fallback):lang(en) { background-color: blue; }
            </style>
            <svg><a id="vector" lang="en">link</a></svg><p class="fallback" lang="en">text</p>
            """);
        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.Document.QuerySelector("#vector")!].GetValue("color"));
        Assert.Equal("provider \"Vector\"", styles[document.Document.QuerySelector("#vector")!].GetValue("string-set"));
        Assert.Equal("rgba(0, 0, 255, 1)", styles[document.Document.QuerySelector(".fallback")!].GetValue("color"));
        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.Document.QuerySelector(".fallback")!].GetValue("background-color"));
    }

    [Fact]
    public void RawRetainedDeclarationsDiscardInvalidSelectorListsAtomically() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>p, :bogus( { string-set: marker 'bad'; }</style><p id='target'>x</p>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#target")!];

        Assert.Equal(string.Empty, style.GetValue("string-set"));
    }

    [Fact]
    public void SelectorListsAndLogicalNestingHonorCancellationAndResourceBounds() {
        Assert.Equal(nameof(HtmlCssSelectorOptions.MaxSelectors), Assert.Throws<HtmlCssSelectorLimitException>(() =>
            HtmlCssSelectorParser.ParseList("a,b", new HtmlCssSelectorOptions { MaxSelectors = 1 })).LimitName);
        Assert.Equal(nameof(HtmlCssSelectorOptions.MaxNestingDepth), Assert.Throws<HtmlCssSelectorLimitException>(() =>
            HtmlCssSelectorParser.Parse(":is(:not(:where(a)))", new HtmlCssSelectorOptions { MaxNestingDepth = 2 })).LimitName);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlCssSelectorParser.ParseList("a,b", cancellationToken: cancellation.Token));
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList("a,").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList(":hover, .").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList("., :hover").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList("p:has(a), p,").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList("p,, :hover").Status);
        Assert.True(HtmlCssSelectorParser.Parse(":is(.picked, .)").IsSupported);
        Assert.True(HtmlCssSelectorParser.Parse(":where(.picked, .)").IsSupported);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse(":not(.picked, .)").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse(":is(.)").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.Unsupported, HtmlCssSelectorParser.Parse(":lang('en')").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.Unsupported, HtmlCssSelectorParser.Parse(":lang(en, fr)").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.Unsupported, HtmlCssSelectorParser.Parse(":lang(åå)").Status);
    }

    [Fact]
    public void IndependentAdvancedSelectorCorpusMatchesOwnedDomAndClassifiesFallback() {
        AdvancedSelectorCase[] corpus = Read<AdvancedSelectorCase>("css-selector-advanced-corpus.json");
        Assert.Equal(35, corpus.Length);
        const string svgNamespace = "http://www.w3.org/2000/svg";
        const string xlinkNamespace = "http://www.w3.org/1999/xlink";
        OfficeIMO.Html.Dom.HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("""
            <html id="root"><body id="body"><main id="main">
              <section id="s1"><p id="p1" class="alpha">one</p><p id="p2" class="beta">two</p></section>
              <section id="s2" lang="en-US"><em id="e1">one</em><p id="p3" class="alpha picked" data-mode="READY">two</p><span id="span-empty"></span></section>
              <section id="s3"><p id="p4">only</p></section>
              <svg id="vector"><a id="va1"></a><a id="va2"></a></svg>
              <a id="ha"></a>
            </main></body></html>
            """).Edit(editor => editor.QuerySelector("#va1")!.SetAttribute(
                new OfficeIMO.Html.Dom.HtmlAttribute("xlink:href", "asset.svg", xlinkNamespace)));
        var options = new HtmlCssSelectorOptions { Namespaces = new HtmlCssNamespaceContext(null,
            new Dictionary<string, string> { ["svg"] = svgNamespace, ["xlink"] = xlinkNamespace }) };
        OfficeIMO.Html.Dom.HtmlElement[] identified = document.QuerySelectorAll("[id]").ToArray();

        foreach (AdvancedSelectorCase item in corpus) {
            HtmlCssSelectorListParseResult result = HtmlCssSelectorParser.ParseList(item.Selector, options);
            Assert.Equal(ParseEnum<HtmlCssSelectorParseStatus>(item.Status), result.Status);
            if (!result.IsSupported) continue;
            string[] actual = identified.Where(result.SelectorList!.Matches).Select(element => element.Id).ToArray();
            Assert.True(item.Matches.SequenceEqual(actual), item.Name + ": expected [" + string.Join(",", item.Matches)
                + "] but matched [" + string.Join(",", actual) + "].");
        }
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
        Assert.False(HtmlCssSelectorParser.Parse("p:empty").Selector!.Matches(
            HtmlDocumentEngine.Default.ParseDocument("<p> </p>").QuerySelector("p")!));
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

    [Fact]
    public void StructuralAndAttributeScansConsumeTheConversionEvaluationBudget() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 5;
        string comments = string.Concat(Enumerable.Repeat("<!-- retained -->", 12));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>section[data-kind]:empty{color:red}</style><section data-kind='evidence'>" + comments + "</section>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlComputedStyleEngine.Compute(document));
        Assert.Equal(nameof(HtmlConversionLimits.MaxSelectorEvaluations), exception.LimitSource);
    }

    [Fact]
    public void SiblingPositionsAreCachedAcrossNthMatchesWithinTheConversionBudget() {
        const int siblingCount = 5_000;
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 30_000;
        string siblings = string.Concat(Enumerable.Range(0, siblingCount).Select(index =>
            "<article data-index='" + index + "'></article>"));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>article:nth-child(odd){color:red}</style><main>" + siblings + "</main>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal(siblingCount, document.Document.QuerySelectorAll("article").Count);
        Assert.Equal("rgba(255, 0, 0, 1)", styles[document.Document.QuerySelector("article")!].GetValue("color"));
    }

    [Fact]
    public void SiblingDiscoveryAccountsForInterveningNonElementNodes() {
        var limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 4;
        string comments = string.Concat(Enumerable.Repeat("<!-- retained -->", 20));
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>p:last-child{color:red}</style><main><p>target</p>" + comments + "</main>",
            new HtmlConversionDocumentOptions { Limits = limits, IncludeNormalizedHtml = false });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(document));
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

    private sealed class AdvancedSelectorCase {
        public string Name { get; set; } = string.Empty;
        public string Selector { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string[] Matches { get; set; } = Array.Empty<string>();
    }
}
