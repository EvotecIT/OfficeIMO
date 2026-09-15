using System.Threading;
using System.Text.Json;
using OfficeIMO.Html.Css;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssSyntaxTests {
    [Fact]
    public void QualifiedCorpusIsLosslessAndRecoversAtDeclaredBoundaries() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Html", "Css", "css-syntax-corpus.json");
        CssSyntaxCorpusCase[] corpus = JsonSerializer.Deserialize<CssSyntaxCorpusCase[]>(
            File.ReadAllText(path), new JsonSerializerOptions { PropertyNameCaseInsensitive = true })!;
        Assert.Equal(12, corpus.Length);

        foreach (CssSyntaxCorpusCase item in corpus) {
            HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(item.Source);
            Assert.True(string.Equals(item.Source, sheet.ToCss(), StringComparison.Ordinal), item.Name + " lost source text.");
            Assert.True(sheet.Rules.Count == item.TopLevelRules, item.Name + " returned an unexpected top-level rule count.");
            Assert.True(CountDeclarations(sheet.Contents) == item.Declarations, item.Name + " returned an unexpected declaration count.");
            Assert.True(sheet.Diagnostics.Count == item.Diagnostics, item.Name + " returned an unexpected diagnostic count.");
        }
    }

    [Fact]
    public void PreservesUnknownRulesDeclarationsNestedValuesAndRecoveryOrder() {
        string css = """
            @future-layer experimental;
            .card, widget-box {
              color: red;
              future-property: paint(foo(1, [two, three]));
              --layout: {columns: 2; gap: 12px};
              .child:hover { unknown-child: yes }
              @future-rule (condition) { nested-value: ok; }
              broken declaration;
              background: blue ! /* retained */ IMPORTANT;
            }
            """;

        HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(css);

        Assert.Equal(css, sheet.ToCss());
        Assert.Equal(2, sheet.Rules.Count);
        HtmlCssAtRule unknown = Assert.IsType<HtmlCssAtRule>(sheet.Rules[0]);
        Assert.Equal("future-layer", unknown.Name);
        Assert.True(unknown.IsStatement);

        HtmlCssQualifiedRule rule = Assert.IsType<HtmlCssQualifiedRule>(sheet.Rules[1]);
        Assert.True(rule.Block!.IsClosed);
        Assert.Collection(rule.Contents,
            value => Assert.Equal("color", Assert.IsType<HtmlCssDeclaration>(value).Name),
            value => {
                HtmlCssDeclaration declaration = Assert.IsType<HtmlCssDeclaration>(value);
                Assert.Equal("future-property", declaration.Name);
                HtmlCssFunctionValue paint = Assert.Single(declaration.Values.OfType<HtmlCssFunctionValue>());
                HtmlCssFunctionValue foo = Assert.Single(paint.Values.OfType<HtmlCssFunctionValue>());
                Assert.Contains(foo.Values, component => component is HtmlCssSimpleBlock block && block.OpeningKind == HtmlCssTokenKind.OpenBracket);
            },
            value => {
                HtmlCssDeclaration declaration = Assert.IsType<HtmlCssDeclaration>(value);
                Assert.Equal("--layout", declaration.Name);
                Assert.True(declaration.IsCustomProperty);
                Assert.Contains(declaration.Values, component => component is HtmlCssSimpleBlock block && block.OpeningKind == HtmlCssTokenKind.OpenBrace);
            },
            value => Assert.IsType<HtmlCssQualifiedRule>(value),
            value => Assert.Equal("future-rule", Assert.IsType<HtmlCssAtRule>(value).Name),
            value => Assert.Equal("CSS001", Assert.IsType<HtmlCssInvalidSyntax>(value).DiagnosticCode),
            value => {
                HtmlCssDeclaration declaration = Assert.IsType<HtmlCssDeclaration>(value);
                Assert.Equal("background", declaration.Name);
                Assert.True(declaration.IsImportant);
            });
        Assert.Single(sheet.Diagnostics);
        Assert.Equal("broken declaration;", sheet.Diagnostics[0].Span.GetText(css).TrimStart());
    }

    [Fact]
    public void RetainsUnclosedBlocksAndFunctionsWithLocations() {
        string css = "@media screen {\r\n  a { color: rgb(1, 2";

        HtmlCssStyleSheet sheet = HtmlCssSyntaxParser.ParseStyleSheet(css);

        Assert.Equal(css, sheet.ToString());
        Assert.Contains(sheet.Diagnostics, diagnostic => diagnostic.Code == "CSS003");
        Assert.Contains(sheet.Diagnostics, diagnostic => diagnostic.Code == "CSS004");
        HtmlCssSourcePosition color = sheet.GetPosition(css.IndexOf("color", StringComparison.Ordinal));
        Assert.Equal(2, color.Line);
        Assert.Equal(7, color.Column);
        Assert.False(Assert.IsType<HtmlCssAtRule>(sheet.Rules[0]).Block!.IsClosed);
    }

    [Fact]
    public void PreservesCommentsAndDuplicateDeclarationOrder() {
        string css = "x{/*a*/color:red;color:future;color:blue}";
        HtmlCssQualifiedRule rule = Assert.IsType<HtmlCssQualifiedRule>(Assert.Single(HtmlCssSyntaxParser.ParseStyleSheet(css).Rules));

        HtmlCssDeclaration[] declarations = rule.Contents.OfType<HtmlCssDeclaration>().ToArray();
        Assert.Equal(new[] { "red", "future", "blue" }, declarations.Select(value => value.ValueSpan.GetText(css)).ToArray());
        Assert.Equal("/*a*/", rule.Block!.Values[0].GetText());
    }

    [Fact]
    public void EmptyDeclarationValueStartsAfterTheColon() {
        const string css = "a{future:}";
        HtmlCssDeclaration declaration = Assert.IsType<HtmlCssDeclaration>(
            Assert.Single(Assert.IsType<HtmlCssQualifiedRule>(Assert.Single(HtmlCssSyntaxParser.ParseStyleSheet(css).Rules)).Contents));

        Assert.Equal(9, declaration.ValueSpan.Offset);
        Assert.Equal(0, declaration.ValueSpan.Length);
    }

    [Fact]
    public void ParsesInlineStyleBlocksWithoutInventingAnOuterRule() {
        const string css = "color:red;future:fn(one[two]);broken value;background:blue";

        HtmlCssStyleBlock block = HtmlCssSyntaxParser.ParseStyleBlock(css);

        Assert.Equal(css, block.ToCss());
        Assert.Equal(new[] { "color", "future", "background" }, block.Declarations.Select(value => value.Name).ToArray());
        Assert.IsType<HtmlCssInvalidSyntax>(block.Contents[2]);
        Assert.Single(block.Diagnostics);
    }

    [Fact]
    public void EnforcesAtomicLimitsAndCancellation() {
        Assert.Equal(nameof(HtmlCssSyntaxOptions.MaxInputCharacters), Assert.Throws<HtmlCssSyntaxLimitException>(() =>
            HtmlCssSyntaxParser.ParseStyleSheet("abc", new HtmlCssSyntaxOptions { MaxInputCharacters = 2 })).LimitName);
        Assert.Equal(nameof(HtmlCssSyntaxOptions.MaxNestingDepth), Assert.Throws<HtmlCssSyntaxLimitException>(() =>
            HtmlCssSyntaxParser.ParseStyleSheet("x{a:f(g(h(i)))}", new HtmlCssSyntaxOptions { MaxNestingDepth = 2 })).LimitName);
        Assert.Equal(nameof(HtmlCssSyntaxOptions.MaxSyntaxNodes), Assert.Throws<HtmlCssSyntaxLimitException>(() =>
            HtmlCssSyntaxParser.ParseStyleSheet("x{a:b;c:d}", new HtmlCssSyntaxOptions { MaxSyntaxNodes = 3 })).LimitName);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => HtmlCssSyntaxParser.ParseStyleSheet("x{a:b}", cancellationToken: cancellation.Token));
    }

    private static int CountDeclarations(IEnumerable<HtmlCssSyntaxNode> nodes) {
        int count = 0;
        foreach (HtmlCssSyntaxNode node in nodes) {
            if (node is HtmlCssDeclaration) count++;
            if (node is HtmlCssRule rule) count += CountDeclarations(rule.Contents);
        }
        return count;
    }

    private sealed class CssSyntaxCorpusCase {
        public string Name { get; set; } = string.Empty;
        public string Source { get; set; } = string.Empty;
        public int TopLevelRules { get; set; }
        public int Declarations { get; set; }
        public int Diagnostics { get; set; }
    }
}
