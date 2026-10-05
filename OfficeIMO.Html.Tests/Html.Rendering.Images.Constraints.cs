using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("width:100px;max-height:20px", 100D, 20D)]
    [InlineData("height:100px;max-width:80px", 80D, 100D)]
    [InlineData("max-width:80px;min-height:60px", 80D, 60D)]
    [InlineData("min-width:100px;max-width:80px", 100D, 50D)]
    public void HtmlImage_AxisConstraintsPreserveDefiniteSizeAndMinimumPrecedence(
        string sizing, double expectedWidth, double expectedHeight) {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(400, 200));
        string html = "<style>body{margin:0}</style><img style='display:block;" + sizing
            + "' src='data:image/png;base64," + data + "'><p>After image</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 600D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(expectedWidth, image.Width, 2);
        Assert.Equal(expectedHeight, image.Height, 2);
        HtmlRenderText[] after = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToArray();
        Assert.Contains("After image", string.Concat(after.Select(text => text.Text)), StringComparison.Ordinal);
        Assert.All(after, text => Assert.True(text.Y >= image.Height));
    }
}
