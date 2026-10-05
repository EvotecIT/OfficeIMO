using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void TransformedTextPaint_KeepsTextMadeVisibleByItsEnclosingEffect() {
        const string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial}</style>"
            + "<div style='position:relative;top:-80px;transform:translateY(80px);transform-origin:0 0'>VisibleEffectMarker</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Contains(Enumerable.Range(0, 24), y => Enumerable.Range(0, 200)
            .Any(x => raster.GetPixel(x, y) != OfficeColor.White));
    }

    [Fact]
    public void TransformedTextPaint_KeepsOverflowGlyphPaintTranslatedIntoThePage() {
        string html = "<style>@page{size:400px 400px;margin:0}body{margin:0;font:16px Arial}</style>"
            + "<div style='position:relative;top:10px;width:200px;white-space:nowrap;transform:translateX(-200px)'>"
            + string.Concat(Enumerable.Repeat("ABCDEFGHIJKLMNO ", 8)) + "</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Contains(Enumerable.Range(300, 100), x => Enumerable.Range(10, 24)
            .Any(y => raster.GetPixel(x, y) != OfficeColor.White));
    }

    [Fact]
    public void TransformedTextPaint_ClipsPositionedOverlaySnapshotsToTheViewport() {
        const string html = "<style>body{margin:0;font:16px Arial}</style>"
            + "<div style='position:absolute;top:-100px'>AboveViewport</div>"
            + "<div style='position:absolute;left:-20px;top:20px'>PartiallyVisible</div>"
            + "<div style='position:absolute;top:700px'>BelowViewport</div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 200D, ViewportHeight = 100D, Margins = HtmlRenderMargins.All(0D)
        });
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        Assert.Contains(Enumerable.Range(20, 24), y => Enumerable.Range(0, 140)
            .Any(x => raster.GetPixel(x, y) != OfficeColor.White));
    }
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
