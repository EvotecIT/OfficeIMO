using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPreformattedWrappingTests {
    [Theory]
    [InlineData("pre-wrap")]
    [InlineData("break-spaces")]
    [InlineData("normal")]
    public void ExplicitWrappingOverridesPreElementDefault(string whiteSpace) {
        var text = Render("white-space:" + whiteSpace + ";overflow-wrap:anywhere");
        Assert.Equal("PublicationChapterIdentifier", string.Concat(text.Select(x => x.Text)));
        Assert.True(text.Select(x => x.Y).Distinct().Count() > 1);
        Assert.All(text, x => Assert.True(x.TextAdvanceWidth <= 60.01D));
    }

    [Theory]
    [InlineData("")]
    [InlineData("white-space:pre;overflow-wrap:anywhere")]
    [InlineData("white-space:nowrap;overflow-wrap:anywhere")]
    public void NonWrappingPreRetainsItsUnbrokenText(string style) {
        var text = Assert.Single(Render(style));
        Assert.Equal("PublicationChapterIdentifier", text.Text);
        Assert.True(text.TextAdvanceWidth > 60D);
    }

    [Fact]
    public void ExplicitNormalCollapsesSpacesWhilePreWrapPreservesThem() {
        var normal = Render("white-space:normal;width:200px", "One  Two");
        var preserved = Render("white-space:pre-wrap;width:200px", "One  Two");
        Assert.Equal("One Two", string.Concat(normal.Select(x => x.Text)));
        Assert.Equal("One  Two", string.Concat(preserved.Select(x => x.Text)));
    }

    private static HtmlRenderText[] Render(string style, string content = "PublicationChapterIdentifier") {
        var result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse(
            "<pre style='width:60px;" + style + "'><code>" + content + "</code></pre>"),
            new HtmlRenderOptions { Mode = HtmlRenderMode.Continuous, ViewportWidth = 120D, Margins = HtmlRenderMargins.All(0D) });
        return result.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
    }
}
