using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(0)]
    [InlineData(40)]
    public void HtmlColumns_PagedOverflowContinuesInsidePageBoundsAtAuthoredTextSize(int precedingHeight) {
        string html = "<style>@page{size:240px 180px;margin:10px}body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<div style='height:" + precedingHeight + "px'></div>"
            + "<section style='height:60px;column-count:2;column-gap:20px;column-fill:auto;column-rule:1px solid blue'>"
            + string.Concat(Enumerable.Range(0, 12).Select(index =>
                "<p><a href='https://example.test/column/" + index + "'>Line" + index.ToString("D2") + "</a></p>"))
            + "</section><p>AfterColumns</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        Assert.True(rendered.Pages.Count >= 2);
        for (int index = 0; index < 12; index++) {
            HtmlRenderText text = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
                item => item.Text == "Line" + index.ToString("D2"));
            Assert.Equal("https://example.test/column/" + index, text.LinkUri);
            Assert.Equal(12D, text.Font.Size, 3);
        }
        Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), item => item.Text == "AfterColumns");
        Assert.All(rendered.Pages, page => {
            Assert.All(page.Visuals.OfType<HtmlRenderText>(), text => {
                Assert.True(text.X >= 10D - 0.001D && text.X + text.Width <= page.Width - 10D + 0.001D);
                Assert.True(text.Y >= 10D - 0.001D && text.Y + text.Height <= page.Height - 10D + 0.001D);
            });
        });
    }

    [Fact]
    public void HtmlColumns_ContinuousFixedHeightKeepsInlineOverflowColumns() {
        string html = "<style>body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<section style='height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + string.Concat(Enumerable.Range(0, 12).Select(index => "<p>Line" + index.ToString("D2") + "</p>"))
            + "</section><p>AfterColumns</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 220D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText[] text = page.Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Equal(13, text.Length);
        Assert.True(Assert.Single(text, item => item.Text == "Line06").X >= 220D);
        Assert.Equal(60D, Assert.Single(text, item => item.Text == "AfterColumns").Y, 3);
    }
    [Theory]
    [InlineData("hidden")]
    [InlineData("clip")]
    public void HtmlColumns_PagedAuthoredClipRetainsTheFixedHeight(string overflow) {
        string html = "<style>@page{size:240px 180px;margin:10px}body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<section style='height:60px;overflow:" + overflow + ";column-count:2;column-gap:20px;column-fill:auto'>"
            + string.Concat(Enumerable.Range(0, 12).Select(index => "<p>Line" + index.ToString("D2") + "</p>"))
            + "</section><p>AfterColumns</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true
        });
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderClipGroup clip = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>());
        Assert.Equal(220D, clip.Width, 3);
        Assert.Equal(60D, clip.Height, 3);
        Assert.Equal(70D, Assert.Single(page.Visuals.OfType<HtmlRenderText>(), item => item.Text == "AfterColumns").Y, 3);
    }

    [Fact]
    public void HtmlColumns_PagedOverflowKeepsAnOverHeightAtomicImageWhole() {
        const string pixel = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=";
        string html = "<style>@page{size:240px 180px;margin:10px}body{margin:0;font:12px/20px Arial}p{margin:0}</style>"
            + "<section style='height:60px;column-count:2;column-gap:20px;column-fill:auto'>"
            + "<img id='atomic-column-image' style='display:block;width:50px;height:80px' src='data:image/png;base64," + pixel + "'>"
            + string.Concat(Enumerable.Range(0, 6).Select(index => "<p>Line" + index.ToString("D2") + "</p>"))
            + "</section><p>AfterColumns</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true,
            ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile()
        });
        HtmlRenderImage image = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(80D, image.Height, 3);
        Assert.Equal(50D, image.Width, 3);
        Assert.True(image.Y >= 10D && image.Y + image.Height <= 170D);
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        Assert.Equal(7, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().Count());
    }

}
