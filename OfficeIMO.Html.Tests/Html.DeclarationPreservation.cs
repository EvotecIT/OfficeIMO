using System.Threading.Tasks;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlDeclarationPreservationTests {
    [Theory]
    [InlineData("width", false)]
    [InlineData("height", false)]
    [InlineData("min-width", false)]
    [InlineData("max-width", false)]
    [InlineData("min-height", false)]
    [InlineData("max-height", false)]
    [InlineData("font", false)]
    [InlineData("border", false)]
    [InlineData("gap", false)]
    [InlineData("width", true)]
    public void ComputedBackgroundPreservesDeclarationLikeUrlParameters(string name, bool quoted) {
        string uri = "https://example.test/pixel;" + name + ":100";
        string url = quoted ? "url('" + uri + "')" : "url(" + uri + ")";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#box{background-image:" + url + ";width:20px;height:20px}</style><div id='box'></div>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#box")!];

        Assert.Contains(uri, style.GetValue("background-image"), StringComparison.Ordinal);
        Assert.Equal("20px", style.GetValue("width"));
    }

    [Theory]
    [InlineData("fn(width:100;height:200;font:bold)")]
    [InlineData("{width:100;height:200;font:bold}")]
    public void CustomPropertyRetainsUrlAndNestedComponents(string components) {
        const string uri = "https://example.test/pixel;min-width:100";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#box{--image:url(" + uri + ");--future:" + components + ";"
            + "background-image:var(--image);min-width:max-content}</style><div id='box'></div>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#box")!];

        Assert.Contains(uri, style.GetValue("background-image"), StringComparison.Ordinal);
        Assert.Equal(components, style.GetValue("--future"));
        Assert.Equal("max-content", style.GetValue("min-width"));
    }

    [Fact]
    public void ProviderFallbackPreservesCustomPropertyCaseAndPriority() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#box{--Width:30px;--width:40px!important;--width:50px;"
            + "--sequence:fn(width:100;height:200)!important;--sequence:later;"
            + "width:var(--Width);height:var(--width);background-image:none}</style><div id='box'></div>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#box")!];

        Assert.Equal("30px", style.GetValue("width"));
        Assert.Equal("40px", style.GetValue("height"));
        Assert.Equal("fn(width:100;height:200)", style.GetValue("--sequence"));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void CommentLikeUrlTextSurvivesDirectAndCustomDeclarations(bool custom, bool quoted) {
        const string uri = "https://example.test/a/*b*/c.png";
        string url = quoted ? "url('" + uri + "')" : "url(" + uri + ")";
        string declarations = custom ? "--image:" + url + ";background-image:var(--image)" : "background-image:" + url;
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<style>#box{" + declarations + ";background-repeat:no-repeat;"
            + "min-width:max-content;/* real comment */width:20px}</style><div id='box'></div>");

        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(document)[document.Document.QuerySelector("#box")!];

        Assert.Contains(uri, style.GetValue("background-image"), StringComparison.Ordinal);
        if (custom) Assert.Contains(uri, style.GetValue("--image"), StringComparison.Ordinal);
        Assert.Equal("20px", style.GetValue("width"));
    }

    [Theory]
    [InlineData("https://example.test/pixel;width:100", false)]
    [InlineData("https://example.test/pixel;width:100", true)]
    [InlineData("https://example.test/a/*b*/c.png", false)]
    [InlineData("https://example.test/a/*b*/c.png", true)]
    public async Task RenderUsesOriginalBackgroundResource(string uri, bool quoted) {
        string url = quoted ? "url('" + uri + "')" : "url(" + uri + ")";
        var requested = new List<string>();
        byte[] png = PdfPngTestImages.CreateRgbPng(2, 2);
        var options = new HtmlRenderOptions {
            ResourceResolver = (request, _) => {
                requested.Add(request.Uri.AbsoluteUri);
                return Task.FromResult<HtmlResolvedResource?>(new HtmlResolvedResource(png, "image/png"));
            },
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            Margins = HtmlRenderMargins.All(0D)
        };

        HtmlRenderDocument rendered = await HtmlRenderTestDriver.RenderAsync(
            "<style>#box{background-image:" + url + ";background-repeat:no-repeat;width:20px;height:20px}</style><div id='box'></div>", options);

        Assert.Equal(new[] { uri }, requested);
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderImage>());
        Assert.False(rendered.HasLoss);
    }
}
