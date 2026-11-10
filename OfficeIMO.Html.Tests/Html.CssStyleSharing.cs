using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlCssStyleSharingTests {
    [Fact]
    public void SiblingsWithEqualTagsAndAttributesKeepPositionalSelectors() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("""
            <style>td { padding: 2px; } td:first-child { color: red; }</style>
            <table><tr><td>A</td><td>B</td><td>C</td></tr></table>
            """);

        var styles = HtmlComputedStyleEngine.Compute(document);
        var cells = document.Document.QuerySelectorAll("td");

        Assert.Equal("rgba(255, 0, 0, 1)", styles[cells[0]].GetValue("color"));
        Assert.NotEqual("rgba(255, 0, 0, 1)", styles[cells[1]].GetValue("color"));
        Assert.Equal(styles[cells[1]].GetValue("color"), styles[cells[2]].GetValue("color"));
    }
}
