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
        Assert.Equal(36, corpus.Length);

        foreach (CssPropertyCorpusCase item in corpus) {
            HtmlCssPropertyParseResult parsed = HtmlCssPropertyParser.Parse(item.Property, item.Value);
            Assert.Equal((HtmlCssPropertyParseStatus)Enum.Parse(typeof(HtmlCssPropertyParseStatus), item.Status), parsed.Status);
            if (item.Kind != null) Assert.Equal((HtmlCssPropertyValueKind)Enum.Parse(typeof(HtmlCssPropertyValueKind), item.Kind), parsed.Value!.Kind);
            if (item.Canonical != null) Assert.Equal(item.Canonical, parsed.Value!.CanonicalText);
        }
    }

    [Fact]
    public void CatalogExposesInheritedAndInitialPropertyContracts() {
        Assert.False(HtmlCssPropertyCatalog.All.Single(property => property.Name == "display").IsInherited);
        Assert.True(HtmlCssPropertyCatalog.All.Single(property => property.Name == "color").IsInherited);
        Assert.Equal("visible", HtmlCssPropertyCatalog.All.Single(property => property.Name == "visibility").InitialValue);
        Assert.Equal("CanvasText", HtmlCssPropertyCatalog.All.Single(property => property.Name == "color").InitialValue);
        Assert.True(HtmlCssPropertyCatalog.All.Single(property => property.Name == "text-transform").IsInherited);
        Assert.Equal("none", HtmlCssPropertyCatalog.All.Single(property => property.Name == "text-transform").InitialValue);
    }

    [Theory]
    [InlineData("display", "GRID", HtmlCssPropertyValueKind.Keyword, "grid")]
    [InlineData("text-transform", "MATH-AUTO", HtmlCssPropertyValueKind.Keyword, "math-auto")]
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
        Assert.Throws<HtmlCssTokenizationLimitException>(() =>
            HtmlCssPropertyParser.Parse("display", "    grid", new HtmlCssTokenizationOptions { MaxInputCharacters = 4 }));
        HtmlCssDeclaration declaration = Assert.Single(
            HtmlCssSyntaxParser.ParseStyleBlock("display:    grid!important;").Declarations);
        Assert.Throws<HtmlCssTokenizationLimitException>(() =>
            HtmlCssPropertyParser.Parse(declaration, new HtmlCssTokenizationOptions { MaxInputCharacters = 4 }));
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
