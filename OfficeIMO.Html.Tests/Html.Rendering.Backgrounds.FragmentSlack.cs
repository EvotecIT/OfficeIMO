using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("stretch")]
    [InlineData("flex-start")]
    public void HtmlRender_PagedAutoHeightFlexItemBackgroundCoversSlack(string alignment) {
        string rows = string.Concat(Enumerable.Range(1, 17).Select(index => $"<p>Footer row {index:00}</p>"));
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;font:16px Arial;line-height:20px}"
            + "*{box-sizing:border-box}p{margin:0}.prelude{height:120px;background:#ddd}"
            + ".row{display:flex;align-items:" + alignment + "}footer{width:90%;background:black;color:white;padding:20px}"
            + ".logo{height:48px;background:#444}.columns{break-inside:avoid;margin-top:24px}.column{width:50%}"
            + ".after{height:40px;background:#eee;color:black}</style><div class='prelude'>Leading label</div>"
            + "<div class='row'><footer><div class='logo'>Logo label</div><div class='columns'><div class='column'>"
            + rows + "</div><div class='column'>Side label</div></div></footer><div></div></div>"
            + "<div class='after'>Trailing label</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, HonorCssPageRules = true, Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(3, rendered.Pages.Count);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "Leading label", "Logo label", "Side label", "Trailing label" })
            Assert.Single(text, item => item.Text == label);
        for (int index = 1; index <= 17; index++)
            Assert.Single(text, item => item.Text == $"Footer row {index:00}");
        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[2].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.Black, first.GetPixel(300, 399));
        Assert.Equal(OfficeColor.White, first.GetPixel(380, 399));
        Assert.Equal(OfficeColor.White, last.GetPixel(300, 399));
    }

    [Theory]
    [InlineData("visible", 0, false)]
    [InlineData("hidden", 0, false)]
    [InlineData("visible", 20, false)]
    [InlineData("hidden", 0, true)]
    public void HtmlRender_PagedBoxBackgroundCoversSlackBeforeKeptChild(string overflow, int pageMargin, bool nested) {
        string rows = string.Concat(Enumerable.Range(1, 13).Select(index => $"<p>Footer row {index:00}</p>"));
        string html = "<style>@page{size:400px 400px;margin:" + pageMargin + "px}html,body{margin:0;font:16px Arial;line-height:20px}"
            + "*{box-sizing:border-box}p{margin:0}.prelude{height:240px;background:#ddd}"
            + "footer{background:black;color:white;padding:20px;overflow:" + overflow + "}"
            + ".logo{height:48px;background:#444}.columns{display:flex;break-inside:avoid;margin-top:24px}"
            + ".column{width:50%}.after{height:40px;background:#eee;color:black}</style>"
            + "<div class='prelude'>Leading label</div>" + (nested ? "<div>" : "") + "<footer><div class='logo'>Logo label</div>"
            + "<div class='columns'><div class='column'>" + rows + "</div><div class='column'>Side label</div>"
            + "</div></footer>" + (nested ? "</div>" : "") + "<div class='after'>Trailing label</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            HonorCssPageRules = true,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Equal(2, rendered.Pages.Count);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        foreach (string label in new[] { "Leading label", "Logo label", "Side label", "Trailing label" })
            Assert.Single(text, item => item.Text == label);
        for (int index = 1; index <= 13; index++)
            Assert.Single(text, item => item.Text == $"Footer row {index:00}");

        OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage last = OfficeDrawingRasterRenderer.Render(rendered.Pages[1].CreateDrawing(), 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.FromRgb(0x44, 0x44, 0x44), first.GetPixel(100, 300));
        Assert.Equal(OfficeColor.Black, first.GetPixel(380 - pageMargin, 350));
        Assert.Equal(OfficeColor.Black, first.GetPixel(380 - pageMargin, 399 - pageMargin));
        Assert.Equal(OfficeColor.White, last.GetPixel(380, 399));
        if (pageMargin > 0) Assert.Equal(OfficeColor.White, first.GetPixel(399, 350));
    }
}
