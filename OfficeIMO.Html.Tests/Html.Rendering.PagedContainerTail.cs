using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 900, 6)]
    [InlineData(true, 900, 6)]
    [InlineData(false, 100, 3)]
    [InlineData(true, 100, 3)]
    public void PagedRendererKeepsLastParagraphWithContainerPadding(bool flex, int imageHeight, int pageCount) {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<style>@page{size:400px 400px;margin:0}html{background:#222}body{margin:0;background:white}"
            + "p{margin:0;font:13px/20px Arial}p:last-child{margin-bottom:16px}"
            + ".primary{padding-bottom:40px}.row{" + (flex ? "display:flex;flex-wrap:wrap" : "display:block")
            + "}.primary,.side{flex:0 0 100%}figure{margin:0}figcaption{font:13px/20px Arial}</style>"
            + "<div class='row'><div class='primary'>"
            + string.Concat(Enumerable.Range(1, 38).Select(index => $"<p>Leading line {index:00}</p>"))
            + "</div><div class='side'><figure><img src='data:image/png;base64," + image
            + "' style='display:block;width:300px;height:" + imageHeight + "px'>"
            + "<figcaption>Figure end</figcaption></figure></div></div><p>Following text</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(pageCount, rendered.Pages.Count);
        Assert.Contains(EnumerateRenderVisuals(rendered.Pages[2].Scene).OfType<HtmlRenderText>(),
            text => text.Text == "Leading line 38");
        Assert.All(rendered.Pages, page => Assert.Contains(EnumerateRenderVisuals(page.Scene),
            visual => visual is HtmlRenderText or HtmlRenderImage));
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 38; index++) Assert.Single(text, item => item.Text == $"Leading line {index:00}");
        Assert.Single(text, item => item.Text == "Figure end");
        Assert.Single(text, item => item.Text == "Following text");
    }
}
