using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlBodyBoxTests {
    [Theory]
    [InlineData("content-box")]
    [InlineData("border-box")]
    public void IntrinsicBodyWidthRetainsTrailingPaddingInTheScrollSurface(string sizing) {
        string html = "<style>*{margin:0}body{width:max-content;box-sizing:" + sizing
            + ";padding:0 20px;background:red}span{display:inline-block;width:300px;height:20px}</style><span></span>";
        HtmlRenderDocument result = Render(html);
        Assert.Equal(340D, result.Pages[0].Width, 3);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Red, raster.GetPixel(339, 10));
        result.RequireNoLoss();
    }

    [Fact]
    public void BodyBorderPaddingAndMarginsParticipateInLayoutAndPaint() {
        var result = Render("<style>*{margin:0}body{margin:5px;padding:7px;border:3px solid blue;font-size:16px;line-height:20px}</style><p>Body</p>");
        var text = Assert.Single(result.Pages[0].Visuals.OfType<HtmlRenderText>(), x => x.Text == "Body");
        Assert.Equal(15D, text.X, 3);
        Assert.Equal(15D, text.Y, 3);
        var image = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, image.GetPixel(5, 20));
        Assert.Equal(OfficeColor.White, image.GetPixel(2, 20));
    }

    [Fact]
    public void PropagatedBodyBackgroundIsNotPaintedTwiceInsideItsBox() {
        var result = Render("<style>*{margin:0}body{margin:10px;padding:10px;border:2px solid blue;background:rgba(255,0,0,.5)}</style><p style='height:20px'></p>");
        var image = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(image.GetPixel(1, 1), image.GetPixel(25, 25));
        Assert.Equal(OfficeColor.Blue, image.GetPixel(10, 25));
    }

    [Fact]
    public void HtmlCanvasBackgroundLeavesTheBodyBackgroundOnItsOwnBox() {
        var result = Render("<style>*{margin:0}html{background:blue}body{margin:10px;padding:10px;border:2px solid black;background:red}</style><p style='height:20px'></p>");
        var image = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, image.GetPixel(1, 1));
        Assert.Equal(OfficeColor.Red, image.GetPixel(25, 25));
        Assert.Equal(OfficeColor.Black, image.GetPixel(10, 25));
    }

    [Fact]
    public void BodyBoxKeepsForcedPageBreaksAndContentOrder() {
        var result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse("<style>*{margin:0}body{padding:5px;border:2px solid blue}p{font-size:16px;line-height:20px}</style><p>First</p><p style='break-before:page'>Second</p>"), new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged, PageSize = new OfficePageSize(200D / 96D, 100D / 96D), Margins = HtmlRenderMargins.All(0D)
        });
        Assert.Equal(2, result.Pages.Count);
        Assert.Contains(result.Pages[0].Visuals.OfType<HtmlRenderText>(), x => x.Text == "First");
        Assert.Contains(result.Pages[1].Visuals.OfType<HtmlRenderText>(), x => x.Text == "Second");
        Assert.DoesNotContain(result.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void NamedPageGeometryKeepsHtmlCanvasBackground() {
        var result = HtmlRenderEngine.Render(HtmlConversionDocument.Parse("<style>@page{size:200px 100px;margin:0}@page chapter{size:240px 100px;margin:0}*{margin:0}html{background:blue}body{margin:10px;padding:10px;background:red}p{page:chapter;height:20px}</style><p></p>"), new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, Margins = HtmlRenderMargins.All(0D) });
        Assert.Equal(240D, result.Pages[0].Width, 3);
        var image = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Blue, image.GetPixel(1, 1));
        Assert.Equal(OfficeColor.Red, image.GetPixel(25, 25));
    }

    [Theory]
    [InlineData("static", 0D)]
    [InlineData("relative", 13D)]
    public void BodyPositionDeterminesAbsoluteContainingBlock(string position, double expected) {
        var result = Render("<style>*{margin:0}body{position:" + position + ";margin:10px;padding:7px;border:3px solid blue}p{height:40px}span{position:absolute;left:0;top:0;font-size:16px;line-height:20px}</style><p></p><span>Absolute</span>");
        var text = Assert.Single(result.Pages[0].Visuals.OfType<HtmlRenderText>(), x => x.Text == "Absolute");
        Assert.Equal(expected, text.X, 3);
        Assert.Equal(expected, text.Y, 3);
    }

    [Theory]
    [InlineData("static", "absolute")]
    [InlineData("relative", "absolute")]
    [InlineData("static", "fixed")]
    [InlineData("relative", "fixed")]
    public void AutomaticPositionedInsetsKeepBodyContentOrigin(string bodyPosition, string childPosition) {
        var result = Render("<style>*{margin:0}body{position:" + bodyPosition + ";margin:10px;padding:7px;border:3px solid blue}span{position:" + childPosition + ";font-size:16px;line-height:20px}</style><span>Positioned</span><p style='height:40px'></p>");
        var text = Assert.Single(result.Pages[0].Visuals.OfType<HtmlRenderText>(), x => x.Text == "Positioned");
        Assert.Equal(20D, text.X, 3);
        Assert.Equal(20D, text.Y, 3);
    }

    [Theory]
    [InlineData("visible", false)]
    [InlineData("hidden", true)]
    public void BodyClipsLocallyOnlyWhenHtmlOwnsViewportOverflow(string htmlOverflow, bool localClip) {
        var result = Render("<style>*{margin:0}html{overflow:" + htmlOverflow + ";}body{margin:10px;height:20px;overflow:hidden}div{height:60px;background:red}</style><div></div>");
        var image = OfficeDrawingRasterRenderer.Render(result.Pages[0].CreateDrawing());
        Assert.Equal(OfficeColor.Red, image.GetPixel(20, 20));
        Assert.Equal(localClip ? OfficeColor.White : OfficeColor.Red, image.GetPixel(20, 45));
    }

    private static HtmlRenderDocument Render(string html) => HtmlRenderEngine.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
        ViewportWidth = 200D, ViewportHeight = 100D, Margins = HtmlRenderMargins.All(0D)
    });
}
