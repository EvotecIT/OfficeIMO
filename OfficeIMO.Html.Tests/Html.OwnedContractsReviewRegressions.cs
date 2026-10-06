using OfficeIMO.Html;
using OfficeIMO.Html.Css;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlOwnedContractsReviewRegressions {
    [Theory]
    [InlineData("n\\a")]
    [InlineData("2n\\a")]
    [InlineData("n-1\\a")]
    public void AnPlusBRequiresAnExactDecodedIdentifier(string value) {
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse($"i:nth-child({value})").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList($"i, :not(i:nth-last-child({value}))").Status);
    }

    [Theory]
    [InlineData("nth-child")]
    [InlineData("nth-last-child")]
    [InlineData("nth-of-type")]
    [InlineData("nth-last-of-type")]
    public void AnPlusBDoesNotConcatenateKeywordTokens(string function) {
        foreach (string value in new[] { "o dd", "e ven", "o/**/dd", "e/**/ven" }) {
            Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse($"i:{function}({value})").Status);
            Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList($"i, :not(i:{function}({value}))").Status);
        }
        Assert.True(HtmlCssSelectorParser.Parse($"i:{function}( odd )").IsSupported);
        Assert.True(HtmlCssSelectorParser.Parse($"i:{function}( even )").IsSupported);
        Assert.True(HtmlCssSelectorParser.Parse($"i:{function}(+/**/n)").IsSupported);
        Assert.True(HtmlCssSelectorParser.Parse($"i:{function}(2n-/**/1)").IsSupported);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse($"i:{function}(+ n)").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse($"i:{function}(2/**/n)").Status);
    }

    [Theory]
    [InlineData("title")]
    [InlineData("textarea")]
    [InlineData("style")]
    [InlineData("xmp")]
    [InlineData("iframe")]
    [InlineData("noembed")]
    [InlineData("noframes")]
    [InlineData("script")]
    [InlineData("plaintext")]
    public void FragmentTextContextPreservesLiteralSvgAttributeOrder(string name) {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument($"<!doctype html><{name} id='context'></{name}>");
        const string source = "<svg><use href='plain' xlink:href='legacy'></use></svg>";
        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(source, document.QuerySelector("#context")!);
        Assert.Equal(source, fragment.TextContent);
        Assert.Empty(fragment.Children);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectorInputBudgetIncludesSurroundingWhitespace(bool list) {
        var limits = new HtmlCssSelectorOptions { MaxInputCharacters = 1 };
        Assert.Throws<HtmlCssSelectorLimitException>(() => {
            if (list) HtmlCssSelectorParser.ParseList("     i ", limits);
            else HtmlCssSelectorParser.Parse("     i ", limits);
        });
    }

    [Theory]
    [InlineData("i:nth-child(2)", true)]
    [InlineData("i:nth-last-child(1)", true)]
    [InlineData("i:nth-of-type(2)", true)]
    [InlineData("i:nth-last-of-type(1)", true)]
    [InlineData("i:first-child", false)]
    [InlineData("i:only-child", false)]
    [InlineData("i + i", true)]
    [InlineData("i ~ i", true)]
    public void FragmentElementsRetainTheirSiblingPositions(string selector, bool expected) {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("<div id='context'></div>");
        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment("<i id='a'></i><!--between--><i id='b'></i>", document.QuerySelector("#context")!);
        Assert.Equal(expected, HtmlCssSelectorParser.Parse(selector).Selector!.Matches(fragment.QuerySelector("#b")!));
    }

    [Theory]
    [InlineData("nth-child")]
    [InlineData("nth-last-child")]
    [InlineData("nth-of-type")]
    [InlineData("nth-last-of-type")]
    public void AnPlusBDoesNotConcatenateIntegerTokens(string function) {
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.Parse($"i:{function}(1 2)").Status);
        Assert.Equal(HtmlCssSelectorParseStatus.InvalidSyntax, HtmlCssSelectorParser.ParseList($"i, :not(i:{function}(2n+1 2))").Status);
        Assert.True(HtmlCssSelectorParser.Parse($"i:{function}(2n + 12)").IsSupported);
    }

    [Theory]
    [InlineData("rgb(1 2 3 /)")]
    [InlineData("rgba(1 2 3 /)")]
    [InlineData("hsl(30 20% 30% /)")]
    [InlineData("hsla(30 20% 30% /)")]
    [InlineData("hwb(30 20% 30% /)")]
    public void PresentAlphaSeparatorRequiresAComponent(string value) {
        Assert.NotEqual(HtmlCssPropertyParseStatus.Parsed, HtmlCssPropertyParser.Parse("color", value).Status);
    }

    [Theory]
    [InlineData("hsl")]
    [InlineData("hsla")]
    [InlineData("hwb")]
    public void FiniteHueRemainsFiniteAfterUnitConversion(string function) {
        HtmlCssPropertyParseResult result = HtmlCssPropertyParser.Parse("color", $"{function}(1e308turn 100% 50%)");
        if (result.Status == HtmlCssPropertyParseStatus.Parsed) {
            double hue = result.Value!.ColorFunction!.Components[0].Value!.Value;
            Assert.False(double.IsNaN(hue) || double.IsInfinity(hue));
            Assert.InRange(hue, 0D, 360D);
        }
    }

    [Theory]
    [InlineData("textarea")]
    [InlineData("title")]
    [InlineData("style")]
    [InlineData("script")]
    public void FragmentTextContextDoesNotInsertAContextStartTag(string name) {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument($"<!doctype html><form><{name} id='context'></{name}></form>");
        HtmlElement context = document.QuerySelector("#context")!;
        string source = $"\nvalue</{name}><b>x</b>";
        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment(source, context);
        Assert.Equal(source, fragment.TextContent);
        Assert.Empty(fragment.Children);
    }

    [Fact]
    public void FragmentRetainsQuirksModeTreeConstructionInsideForm() {
        HtmlDocument document = HtmlDocumentEngine.Default.ParseDocument("<form><div id='context'></div></form>");
        Assert.Equal(HtmlDocumentMode.Quirks, document.Mode);
        HtmlDocumentFragment fragment = HtmlDocumentEngine.Default.ParseFragment("<p><table><tr><td>x</td></tr></table>", document.QuerySelector("#context")!);
        Assert.NotNull(fragment.QuerySelector("p > table"));
    }
}
