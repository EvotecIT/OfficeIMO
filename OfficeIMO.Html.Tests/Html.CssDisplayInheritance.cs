using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssDisplayInheritanceTests {
    [Theory]
    [InlineData("div", "block")]
    [InlineData("span", "inline")]
    [InlineData("ul", "block")]
    public void ExplicitDisplayInheritUsesTheParentsEffectiveDefaultDisplay(
        string parentTag, string expectedDisplay) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<" + parentTag + "><label id='tile' style='display:inherit'>Choice</label></" + parentTag + ">");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#tile")!];

        Assert.Equal(expectedDisplay, style.GetValue("display"));
    }

    [Fact]
    public void StylesheetDisplayInheritUsesTheParentsEffectiveDefaultDisplay() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#tile { display: inherit; }</style><div><label id='tile'>Choice</label></div>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#tile")!];

        Assert.Equal("block", style.GetValue("display"));
    }

    [Theory]
    [InlineData("initial", "inline")]
    [InlineData("unset", "inline")]
    [InlineData("revert", "block")]
    public void DisplayInheritUsesTheParentsEffectiveDisplayAfterCssWideReset(
        string reset, string expectedDisplay) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<div id='parent' style='display:" + reset + "'><label id='tile' style='display:inherit'>Choice</label></div>");
        var styles = HtmlComputedStyleEngine.Compute(document);

        Assert.Equal(expectedDisplay, styles[document.Document.QuerySelector("#tile")!].GetValue("display"));
        if (reset != "revert") {
            Assert.Equal(expectedDisplay, styles[document.Document.QuerySelector("#parent")!].GetValue("display"));
        }
    }
}
