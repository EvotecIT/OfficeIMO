using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(HtmlRenderMode.Paged)]
    [InlineData(HtmlRenderMode.Continuous)]
    public void HtmlGridDefaultBudgetAllowsIndependentRows(HtmlRenderMode mode) {
        string html = "<style>@page{size:1600px 400px;margin:0}html,body{margin:0}"
            + ".grid{display:grid;grid-template-columns:repeat(16,100px)}"
            + ".grid>div{height:20px;font:12px/20px Arial}</style><div class='grid'>"
            + string.Concat(Enumerable.Range(1, 4096).Select(index => $"<div>Row {index:0000}</div>")) + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = mode, ViewportWidth = 1600D, Margins = HtmlRenderMargins.All(0D) });
        string[] labels = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().Select(item => item.Text).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        Assert.Equal(Enumerable.Range(1, 4096).Select(index => $"Row {index:0000}"), labels);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlGridPagingPreservesLinesInAnOversizedSharedRow(bool paintBackground) {
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0;font:16px Arial,sans-serif;line-height:20px}"
            + "p{margin:0}.row{display:grid;grid-template-columns:90% 10%}footer{padding:20px;"
            + (paintBackground ? "background:black;color:white;" : "") + "}</style>"
            + "<div style='height:120px'>Leading label</div><div class='row'><footer>"
            + "<div style='height:48px'>Logo label</div><div style='margin-top:24px'>"
            + string.Concat(Enumerable.Range(1, 17).Select(index => $"<p>Footer row {index:00}</p>"))
            + "<p>Side label</p></div></footer><div></div></div><div style='height:40px'>Trailing label</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 17; index++) Assert.Single(text, item => item.Text == $"Footer row {index:00}");
        Assert.Single(text, item => item.Text == "Side label");
        Assert.Single(text, item => item.Text == "Trailing label");
        Assert.All(rendered.Pages, page => Assert.All(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderText>(),
            item => Assert.InRange(item.Y + item.Height, 0D, page.Height + 0.001D)));
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
        if (paintBackground) {
            OfficeRasterImage first = OfficeDrawingRasterRenderer.Render(rendered.Pages[0].CreateDrawing());
            Assert.Equal(OfficeColor.Black, first.GetPixel(10, 250));
        }
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlGridPagingPreservesAtomicSiblingImagesAndFittingKeeps(bool hardKeep) {
        string image = Convert.ToBase64String(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<style>@page{size:400px 400px;margin:0}html,body{margin:0}p{margin:0;font:16px/20px Arial}"
            + ".row{display:grid;grid-template-columns:50% 50%;align-items:start}</style>"
            + "<div class='row'><div>"
            + string.Concat(Enumerable.Range(1, 36).Select(index => $"<p>Prose {index:00}</p>"))
            + "</div><div><div style='height:310px'></div><div style='"
            + (hardKeep ? "break-inside:avoid" : "") + "'><p>Image lead</p>"
            + "<img style='display:block;width:100px;height:100px' src='data:image/png;base64," + image
            + "'><p>Image end</p></div></div></div><p>Following text</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();
        for (int index = 1; index <= 36; index++) Assert.Single(text, item => item.Text == $"Prose {index:00}");
        Assert.Single(text, item => item.Text == "Image lead");
        Assert.Single(text, item => item.Text == "Image end");
        Assert.Single(text, item => item.Text == "Following text");
        Assert.Single(rendered.Pages.SelectMany(page => EnumerateRenderVisuals(page.Scene)).OfType<HtmlRenderImage>());
        if (hardKeep) {
            HtmlRenderPage imagePage = Assert.Single(rendered.Pages, page =>
                EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderImage>().Any());
            Assert.Contains(EnumerateRenderVisuals(imagePage.Scene).OfType<HtmlRenderText>(), item => item.Text == "Image lead");
            Assert.Contains(EnumerateRenderVisuals(imagePage.Scene).OfType<HtmlRenderText>(), item => item.Text == "Image end");
        }
    }

}
