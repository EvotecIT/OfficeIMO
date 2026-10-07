using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, 900, 5, 3)]
    [InlineData(false, 100, 3, 1)]
    [InlineData(true, 900, 5, 3)]
    [InlineData(true, 100, 3, 1)]
    public void HtmlFlexGapDoesNotStrandPageBeforeImage(bool column, int imageHeight, int pageCount, int imageFragments) {
        string html = WrappedFlexGapFixture(imageHeight, column: column);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        Assert.Equal(pageCount, rendered.Pages.Count);
        HtmlRenderImage[] images = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderImage>().ToArray();
        Assert.Equal(imageFragments, images.Length);
        Assert.Equal(0D, images[0].Y, precision: 3);
        Assert.All(images, image => Assert.Equal(imageHeight, image.Height, precision: 3));
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 38; index++) Assert.Single(text, item => item.Text == $"Leading line {index:00}");
        Assert.Single(text, item => item.Text == "Figure end");
        Assert.Single(text, item => item.Text == "Following text");
        Assert.All(rendered.Pages, page => Assert.Contains(EnumerateRenderVisuals(page.Scene),
            visual => visual is HtmlRenderText or HtmlRenderImage));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void HtmlFlexGapPreservesContainerBorderPaint(bool column, bool varyingPageWidth) {
        string html = WrappedFlexGapFixture(900, "border-left:1px solid red", column);
        if (varyingPageWidth) html = WithVaryingPageWidth(html);
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });

        HtmlRenderPage gapPage = rendered.Pages.First(page =>
            !EnumerateRenderVisuals(page.Scene).Any(visual => visual is HtmlRenderText or HtmlRenderImage));
        Assert.Contains(EnumerateRenderVisuals(gapPage.Scene).OfType<HtmlRenderShape>(),
            shape => shape.Source == "div.row:border-left" && shape.Shape.StrokeColor == OfficeColor.Red);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlFlexGapPageGeometryChangeDoesNotResumeCompletedContent(bool column) {
        string html = WithVaryingPageWidth(WrappedFlexGapFixture(900, column: column));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 38; index++) Assert.Single(text, item => item.Text == $"Leading line {index:00}");
        Assert.Single(text, item => item.Text == "Figure end");
        Assert.Single(text, item => item.Text == "Following text");
    }

    private static string WithVaryingPageWidth(string html) => html.Replace(
        "@page{size:400px 400px;margin:0}",
        "@page{size:500px 400px;margin:0}@page:first{size:400px 400px}");

    private static string WrappedFlexGapFixture(int imageHeight, string rowStyle = "", bool column = false) {
        string image = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        return "<style>@page{size:400px 400px;margin:0}html{background:#222}body{margin:0;background:white}"
            + "p{margin:0;font:13px/20px Arial}figure{margin:0}figcaption{font:13px/20px Arial}"
            + ".row{display:flex;" + (column ? "flex-direction:column;" : "flex-wrap:wrap;")
            + "gap:56px;" + rowStyle + "}.primary,.side{flex:0 0 " + (column ? "auto" : "100%") + "}</style>"
            + "<div class='row'><div class='primary'>"
            + string.Concat(Enumerable.Range(1, 38).Select(index => $"<p>Leading line {index:00}</p>"))
            + "</div><div class='side'><figure><img src='data:image/png;base64," + image
            + "' style='display:block;width:300px;height:" + imageHeight + "px'><figcaption>Figure end</figcaption>"
            + "</figure></div></div><p>Following text</p>";
    }
}
