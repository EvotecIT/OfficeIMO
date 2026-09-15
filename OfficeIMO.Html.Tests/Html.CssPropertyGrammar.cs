using System.Threading;
using System.Text.Json;
using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using Xunit;

namespace OfficeIMO.Tests;

[Collection(HtmlCssPropertyGrammarCollection.Name)]
public sealed class HtmlCssPropertyGrammarTests {
    [Fact]
    public void IndependentPropertyGrammarCorpusMatchesDeclaredSlice() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", "css-property-grammar-corpus.json");
        CssPropertyCorpusCase[] corpus = JsonSerializer.Deserialize<CssPropertyCorpusCase[]>(
            File.ReadAllText(path), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
        Assert.Equal(22, corpus.Length);

        foreach (CssPropertyCorpusCase item in corpus) {
            HtmlCssPropertyParseResult parsed = HtmlCssPropertyParser.Parse(item.Property, item.Value);
            Assert.Equal((HtmlCssPropertyParseStatus)Enum.Parse(typeof(HtmlCssPropertyParseStatus), item.Status), parsed.Status);
            if (item.Kind != null) Assert.Equal((HtmlCssPropertyValueKind)Enum.Parse(typeof(HtmlCssPropertyValueKind), item.Kind), parsed.Value!.Kind);
            if (item.Canonical != null) Assert.Equal(item.Canonical, parsed.Value!.CanonicalText);
        }
    }

    [Fact]
    public void CatalogExposesTheFirstOwnedPropertyDefinitions() {
        Assert.Equal(new[] { "display", "visibility", "opacity", "color" },
            HtmlCssPropertyCatalog.All.Select(property => property.Name).ToArray());
        Assert.False(HtmlCssPropertyCatalog.All.Single(property => property.Name == "display").IsInherited);
        Assert.True(HtmlCssPropertyCatalog.All.Single(property => property.Name == "color").IsInherited);
        Assert.Equal("visible", HtmlCssPropertyCatalog.All.Single(property => property.Name == "visibility").InitialValue);
        Assert.Equal("CanvasText", HtmlCssPropertyCatalog.All.Single(property => property.Name == "color").InitialValue);
    }

    [Theory]
    [InlineData("display", "GRID", HtmlCssPropertyValueKind.Keyword, "grid")]
    [InlineData("visibility", "collapse", HtmlCssPropertyValueKind.Keyword, "collapse")]
    [InlineData("opacity", ".5", HtmlCssPropertyValueKind.Number, "0.5")]
    [InlineData("opacity", "125%", HtmlCssPropertyValueKind.Percentage, "125%")]
    [InlineData("color", "RebeccaPurple", HtmlCssPropertyValueKind.NamedColor, "rebeccapurple")]
    [InlineData("color", "CanvasText", HtmlCssPropertyValueKind.SystemColor, "canvastext")]
    [InlineData("color", "#F0a8", HtmlCssPropertyValueKind.HexColor, "#f0a8")]
    [InlineData("color", "currentColor", HtmlCssPropertyValueKind.CurrentColor, "currentcolor")]
    public void ParsesTypedValuesWithoutMakingARenderingClaim(
        string property,
        string value,
        HtmlCssPropertyValueKind kind,
        string canonical) {
        HtmlCssPropertyParseResult parsed = HtmlCssPropertyParser.Parse(property, value);

        Assert.Equal(HtmlCssPropertyParseStatus.Parsed, parsed.Status);
        Assert.Equal(kind, parsed.Value!.Kind);
        Assert.Equal(canonical, parsed.Value.CanonicalText);
    }

    [Theory]
    [InlineData("display", "grid inherit")]
    [InlineData("visibility", "force-hidden")]
    [InlineData("opacity", "opaque")]
    [InlineData("color", "not-a-color")]
    [InlineData("color", "#12")]
    public void KeepsKnownButUnimplementedValuesDistinctFromUnknownProperties(string property, string value) {
        HtmlCssPropertyParseResult unsupported = HtmlCssPropertyParser.Parse(property, value);
        HtmlCssPropertyParseResult unknown = HtmlCssPropertyParser.Parse("future-property", value);

        Assert.Equal(HtmlCssPropertyParseStatus.UnsupportedValue, unsupported.Status);
        Assert.Equal(HtmlCssPropertyParseStatus.UnknownProperty, unknown.Status);
        Assert.NotNull(unsupported.Definition);
        Assert.Null(unknown.Definition);
    }

    [Fact]
    public void ParsesCssWideKeywordsAndDefersVarUntilComputedValueTime() {
        HtmlCssPropertyParseResult wide = HtmlCssPropertyParser.Parse("display", "ReVeRt-LaYeR");
        HtmlCssPropertyParseResult deferred = HtmlCssPropertyParser.Parse("color", "var(--accent, red)");

        Assert.Equal(HtmlCssWideKeyword.RevertLayer, wide.Value!.CssWideKeyword);
        Assert.Equal(HtmlCssPropertyParseStatus.Deferred, deferred.Status);
        Assert.Equal(HtmlCssPropertyValueKind.DeferredFunction, deferred.Value!.Kind);
        Assert.Equal(HtmlCssPropertyParseStatus.InvalidSyntax,
            HtmlCssPropertyParser.Parse("color", "var(red)").Status);
    }

    [Fact]
    public void DeclarationParsingRetainsSourceAndSeparatesImportant() {
        HtmlCssDeclaration declaration = Assert.Single(
            HtmlCssSyntaxParser.ParseStyleBlock("color: /* tone */ ReD !/**/ IMPORTANT;").Declarations);

        HtmlCssPropertyParseResult parsed = HtmlCssPropertyParser.Parse(declaration);

        Assert.True(parsed.IsImportant);
        Assert.Equal("/* tone */ ReD", parsed.AuthoredValue);
        Assert.Equal("red", parsed.Value!.CanonicalText);
        Assert.Equal(" /* tone */ ReD !/**/ IMPORTANT", declaration.ValueText);
    }

    [Fact]
    public void MalformedComponentsAreInvalidAndCancellationPublishesNoResult() {
        Assert.Equal(HtmlCssPropertyParseStatus.InvalidSyntax,
            HtmlCssPropertyParser.Parse("color", "var(--accent").Status);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            HtmlCssPropertyParser.Parse("display", "block", cancellationToken: cancellation.Token));
        Assert.Throws<HtmlCssTokenizationLimitException>(() =>
            HtmlCssPropertyParser.Parse("display", "inline-block", new HtmlCssTokenizationOptions { MaxInputCharacters = 4 }));
    }

    [Fact]
    public void ManagedCascadeTraceExplainsSelectorLayerImportantAndInlinePrecedence() {
        const string html = """
            <style>
              @layer base, theme;
              @layer base { #target { color:red !important; display:block; } }
              @layer theme { .item { color:blue !important; display:grid; } }
              #target { visibility:hidden; opacity:.25; }
            </style>
            <p id="target" class="item" style="color:lime !important; display:flex">Trace</p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        OfficeIMO.Html.Dom.HtmlElement target = document.Document.QuerySelector("#target")!;
        Assert.Null(HtmlComputedStyleEngine.Compute(document)[target].GetCascadeTrace("color"));
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document,
            new HtmlComputedStyleOptions { IncludeCascadeTraces = true })[target];

        HtmlCssCascadeTrace color = Assert.IsType<HtmlCssCascadeTrace>(style.GetCascadeTrace("color"));
        Assert.Equal("lime", color.ComputedValue);
        HtmlCssCascadeCandidate selectedColor = Assert.Single(color.Candidates,
            candidate => candidate.Decision == HtmlCssCascadeDecision.Selected);
        Assert.Equal(HtmlCssCascadeSourceKind.InlineStyle, selectedColor.Source);
        Assert.True(selectedColor.IsImportant);
        Assert.Equal(HtmlCssPropertyParseStatus.Parsed, selectedColor.GrammarStatus);
        Assert.Contains(color.Candidates, candidate => candidate.Selector == "#target" && candidate.LayerName == "base");
        Assert.Contains(color.Candidates, candidate => candidate.Selector == ".item" && candidate.LayerName == "theme");

        HtmlCssCascadeTrace display = Assert.IsType<HtmlCssCascadeTrace>(style.GetCascadeTrace("DISPLAY"));
        Assert.Equal("flex", display.ComputedValue);
        Assert.Equal(HtmlCssCascadeSourceKind.InlineStyle,
            Assert.Single(display.Candidates, candidate => candidate.Decision == HtmlCssCascadeDecision.Selected).Source);
        Assert.Equal("hidden", style.GetCascadeTrace("visibility")!.ComputedValue);
        Assert.Equal("0.25", style.GetCascadeTrace("opacity")!.ComputedValue);
        Assert.Null(style.GetCascadeTrace("margin-left"));
    }

    [Fact]
    public void TraceDistinguishesInheritanceResetAndInvalidValueRejection() {
        const string html = """
            <style>
              body { color:red; visibility:hidden; }
              #inherited { color:not-a-color; }
              #reset { visibility:initial; }
            </style>
            <p id="inherited">Inherited</p><p id="reset">Reset</p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(
            document, new HtmlComputedStyleOptions { IncludeCascadeTraces = true });

        HtmlCssCascadeTrace inherited = styles[document.Document.QuerySelector("#inherited")!].GetCascadeTrace("color")!;
        Assert.True(inherited.IsInherited);
        Assert.Equal("rgba(255, 0, 0, 1)", inherited.ComputedValue);
        Assert.Single(inherited.Candidates);
        Assert.Equal(HtmlCssCascadeDecision.Inherited, inherited.Candidates[0].Decision);
        Assert.False(inherited.Candidates[0].IsEffective);

        HtmlCssCascadeTrace reset = styles[document.Document.QuerySelector("#reset")!].GetCascadeTrace("visibility")!;
        Assert.False(reset.IsInherited);
        Assert.True(reset.IsReset);
        Assert.Equal("visible", reset.ComputedValue);
        Assert.Equal("initial", Assert.Single(reset.Candidates).DeclaredValue);
        Assert.Equal(HtmlCssCascadeDecision.Reset, reset.Candidates[0].Decision);
        Assert.True(reset.Candidates[0].IsEffective);
    }

    [Fact]
    public void InvalidCustomPropertySubstitutionUsesInheritedOrInitialFallback() {
        const string html = """
            <style>
              body { color:red; }
              #missing { color:var(--missing); opacity:var(--missing); }
              #cycle { --a:var(--b); --b:var(--a); color:var(--a); opacity:var(--a); }
            </style>
            <body><p id="missing">Missing</p><p id="cycle">Cycle</p></body>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        IReadOnlyDictionary<OfficeIMO.Html.Dom.HtmlElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(
            document, new HtmlComputedStyleOptions { IncludeCascadeTraces = true });

        foreach (string id in new[] { "#missing", "#cycle" }) {
            HtmlComputedStyle style = styles[document.Document.QuerySelector(id)!];
            HtmlCssCascadeTrace color = style.GetCascadeTrace("color")!;
            Assert.Equal("rgba(255, 0, 0, 1)", style.GetValue("color"));
            Assert.True(color.IsInherited);
            Assert.False(color.IsReset);
            Assert.Equal(HtmlCssCascadeDecision.InvalidAtComputedValue, Assert.Single(color.Candidates).Decision);
            Assert.True(color.Candidates[0].IsEffective);

            HtmlCssCascadeTrace opacity = style.GetCascadeTrace("opacity")!;
            Assert.Equal("1", style.GetValue("opacity"));
            Assert.False(opacity.IsInherited);
            Assert.True(opacity.IsReset);
            Assert.Equal(HtmlCssCascadeDecision.InvalidAtComputedValue, Assert.Single(opacity.Candidates).Decision);
            Assert.True(opacity.Candidates[0].IsEffective);
        }
    }

    [Fact]
    public void TraceMarksOverriddenWideKeywordsWithoutClaimingTheyChangedTheResult() {
        const string html = """
            <p id="target" style="color:initial;color:blue;opacity:revert;opacity:.5;display:revert-layer;display:flex">Winner</p>
            """;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(html);
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document,
            new HtmlComputedStyleOptions { IncludeCascadeTraces = true })[document.Document.QuerySelector("#target")!];

        foreach (string property in new[] { "color", "opacity", "display" }) {
            HtmlCssCascadeTrace trace = style.GetCascadeTrace(property)!;
            Assert.Single(trace.Candidates, candidate => candidate.IsEffective);
            Assert.Equal(HtmlCssCascadeDecision.Selected, Assert.Single(trace.Candidates, candidate => candidate.IsEffective).Decision);
        }
        Assert.Equal(HtmlCssCascadeDecision.Overridden,
            Assert.Single(style.GetCascadeTrace("color")!.Candidates, candidate => candidate.DeclaredValue == "initial").Decision);
        Assert.Equal(HtmlCssCascadeDecision.Overridden,
            Assert.Single(style.GetCascadeTrace("opacity")!.Candidates, candidate => candidate.DeclaredValue == "revert").Decision);
        Assert.Equal(HtmlCssCascadeDecision.Overridden,
            Assert.Single(style.GetCascadeTrace("display")!.Candidates, candidate => candidate.DeclaredValue == "revert-layer").Decision);
    }

    [Fact]
    public void InlineSyntaxUsesTheCallersCssLimitsAndHasNoIndependentEightMegabyteCeiling() {
        string largeValue = new string('a', 8 * 1024 * 1024 + 64);
        OfficeIMO.Html.Dom.HtmlDocument unbounded = HtmlDocumentEngine.Default.ParseDocument(
            "<p id='large' style='--payload:" + largeValue + "'>Large</p>");
        Assert.NotNull(HtmlComputedStyleEngine.Compute(unbounded)[unbounded.QuerySelector("#large")!]);

        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxCssBytes = 16;
        limits.MaxTotalCssBytes = 16;
        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.Parse(
            "<p id='bounded' style='color:red;opacity:.5'>Bounded</p>",
            new HtmlConversionDocumentOptions { Limits = limits }));
        Assert.Equal(HtmlConversionDiagnosticCodes.CssSizeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxCssBytes), exception.LimitSource);
    }

    [Fact]
    public void InlineSyntaxTranslatesTheConfiguredTokenLimitToTheConversionContract() {
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxCssTokens = 2;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p id='bounded' style='color:red'>Bounded</p>",
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(document));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssTokenLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxCssTokens), exception.LimitSource);
        Assert.Equal(3, exception.Actual);
        Assert.Equal(2, exception.Limit);
    }

    [Fact]
    public void InlineSyntaxTranslatesTheConfiguredSyntaxNodeLimitToTheConversionContract() {
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxCssTokens = 32;
        limits.MaxCssSyntaxNodes = 1;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p id='bounded' style='color:red'>Bounded</p>",
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(document));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssSyntaxNodeLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxCssSyntaxNodes), exception.LimitSource);
        Assert.Equal(2, exception.Actual);
        Assert.Equal(1, exception.Limit);
    }

    private sealed class CssPropertyCorpusCase {
        public string Name { get; set; } = string.Empty;
        public string Property { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string? Kind { get; set; }
        public string? Canonical { get; set; }
    }
}

[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class HtmlCssPropertyGrammarCollection {
    public const string Name = "HTML CSS property grammar";
}
