using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("j", 80)]
    [InlineData("ffffffffffffffffffffffffffffffffffffffffffffffffff", -200)]
    public void TransformedTextPaint_PreservesScopedItalicGlyphInk(string value, int offset) {
        string face = Convert.ToBase64String(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSansPro-Regular.otf")));
        string css = "<style>@page{size:400px 400px;margin:0}body{margin:0}"
            + "@font-face{font-family:Proof;src:url(data:font/otf;base64," + face + ")}"
            + "div{font:italic 40px/60px Proof;width:20px;white-space:nowrap;transform-origin:0 0}</style>";
        HtmlRenderDocument control = HtmlRenderTestDriver.Render(css + "<div style='position:relative;left:" + offset + "px'>" + value + "</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderDocument actual = HtmlRenderTestDriver.Render(css + "<div style='transform:translateX(" + offset + "px)'>" + value + "</div>",
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(Assert.Single(control.Pages).CreateDrawing());
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(actual.Pages).CreateDrawing());
        int ink = 0;
        for (int y = 0; y < 80; y++) for (int x = 0; x < 400; x++) {
            if (expected.GetPixel(x, y) != OfficeColor.White) ink++;
            Assert.Equal(expected.GetPixel(x, y), raster.GetPixel(x, y));
        }
        Assert.True(ink > 20);
    }

}
