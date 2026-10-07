using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlWidthContainmentTests {
    [Theory]
    [InlineData("width:200px", 200)]
    [InlineData("min-width:200px", 200)]
    [InlineData("width:200px;max-width:80px", 80)]
    [InlineData("width:auto", 100)]
    [InlineData("width:200px;display:flex", 200)]
    [InlineData("width:200px;display:grid", 200)]
    public void ContainingWidthDoesNotActAsAnImplicitMaximum(string sizing, double expectedWidth) {
        string html = "<html><body style='margin:0'><div style='width:100px'><div id='child' style='height:10px;background:#124e80;" + sizing + "'></div></div></body></html>";
        var rendered = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 300, ViewportHeight = 100, Margins = HtmlRenderMargins.All(0)
        });
        var child = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), v => v.Source == "div#child");
        Assert.Equal(expectedWidth, child.Width, 4);
    }
}
