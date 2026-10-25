using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlBackgroundLengthContexts_ImageSizeAndPositionUseTheNearestSizeContainer() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        string html = "<div style='container-type:size;width:200px;height:100px'>"
            + "<div style='width:80px;height:40px;background-image:url(data:image/png;base64," + data
            + ");background-size:10cqw 10cqh;background-position:5cqw 5cqh;background-repeat:no-repeat'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous, ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(20D, image.Width, 6);
        Assert.Equal(10D, image.Height, 6);
        Assert.Equal(10D, image.X, 6);
        Assert.Equal(5D, image.Y, 6);
        rendered.RequireNoLoss();
    }
}
