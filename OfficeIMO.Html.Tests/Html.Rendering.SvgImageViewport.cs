using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlSvgImageViewportTests {
    private static readonly OfficeColor SvgColor = OfficeColor.FromRgb(36, 95, 158);
    private static readonly OfficeColor BoxColor = OfficeColor.FromRgb(238, 238, 238);
    private static readonly OfficeColor FollowingColor = OfficeColor.FromRgb(208, 108, 36);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgWithoutViewBox_ScalesItsIntrinsicCanvasForImagesAndBackgrounds(bool background) {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='128' height='64'>"
            + "<rect width='64' height='64' fill='red'/><rect x='64' width='64' height='64' fill='blue'/></svg>";
        string uri = "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
        string html = background
            ? "<div style='width:96px;height:64px;background-color:#eee;background-image:url(" + uri
                + ");background-size:64px 53.949px;background-repeat:no-repeat'></div>"
            : "<img src='" + uri + "' style='display:block;width:64px;height:53.949px;background:#eee'>";
        HtmlRenderPage page = RenderViewport(html);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        // A source without viewBox scales its intrinsic canvas as an image.
        // The independent browser witness retains both halves across the box.
        foreach (int y in new[] { 20, 32, 52, 65 }) {
            Assert.Equal(OfficeColor.Red, raster.GetPixel(24, y));
            Assert.Equal(OfficeColor.Blue, raster.GetPixel(64, y));
        }
        Assert.Equal(background ? BoxColor : OfficeColor.White, raster.GetPixel(84, 40));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExternalSvgImage_UsesItsConcreteViewportWithoutChangingBoxOrFlow(bool stretch) {
        string html = "<img id='picture' src='" + SvgDataUri(stretch)
            + "' style='display:block;width:64px;height:53.949px;background:#eee'>"
            + FollowingBlock;
        HtmlRenderPage page = RenderViewport(html);
        HtmlRenderDrawing image = Assert.Single(page.Visuals.OfType<HtmlRenderDrawing>());
        HtmlRenderShape following = FollowingVisual(page);

        // Keep the image's CSS box on the used-layout grid. Aspect fitting only
        // changes the SVG paint inside that box, never the following flow.
        Assert.Equal(16D, image.X, 6);
        Assert.Equal(16D, image.Y, 6);
        Assert.Equal(64D, image.Width, 6);
        Assert.Equal(53.9375D, image.Height, 6);
        Assert.Equal(image.Y + image.Height, following.Y, 6);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        AssertSvgPaint(raster, stretch);
        Assert.Equal(OfficeColor.White, raster.GetPixel(84, 40));
        Assert.Equal(FollowingColor, raster.GetPixel(32, 71));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgBackground_UsesItsConcreteTileViewportWithoutChangingBoxOrFlow(bool stretch) {
        string html = "<div id='picture' style='width:96px;height:64px;background-color:#eee;"
            + "background-image:url(" + SvgDataUri(stretch) + ");"
            + "background-size:64px 53.949px;background-repeat:no-repeat;background-position:left top'></div>"
            + FollowingBlock;
        HtmlRenderPage page = RenderViewport(html);
        HtmlRenderShape box = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(),
            visual => visual.Source == "div#picture" && visual.Shape.FillColor == BoxColor);
        HtmlRenderShape following = FollowingVisual(page);

        Assert.Equal(16D, box.X, 6);
        Assert.Equal(16D, box.Y, 6);
        Assert.Equal(96D, box.Width, 6);
        Assert.Equal(64D, box.Height, 6);
        Assert.Equal(box.Y + box.Height, following.Y, 6);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing());

        AssertSvgPaint(raster, stretch);
        Assert.Equal(BoxColor, raster.GetPixel(84, 40));
        Assert.Equal(BoxColor, raster.GetPixel(32, 75));
        Assert.Equal(FollowingColor, raster.GetPixel(32, 82));
    }

    private const string FollowingBlock = "<div id='following' style='width:96px;height:4px;background:#d06c24'></div>";

    private static string SvgDataUri(bool stretch) {
        string aspectRatio = stretch ? " preserveAspectRatio='none'" : string.Empty;
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='128' height='64' viewBox='0 0 128 64'"
            + aspectRatio + "><rect width='128' height='64' fill='#245f9e'/></svg>";
        return "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
    }

    private static HtmlRenderPage RenderViewport(string content) => Assert.Single(HtmlRenderTestDriver.Render(
        "<!doctype html><style>html,body{margin:0;padding:0;background:white}"
        + "body{padding:16px;font-size:0;line-height:0}</style>" + content,
        new HtmlRenderOptions {
            ViewportWidth = 160D,
            ViewportHeight = 120D,
            Margins = HtmlRenderMargins.All(0D),
            UserAgentStyles = HtmlRenderUserAgentStyleMode.Browser,
            AllowSystemFontFallback = false
        }).Pages);

    private static HtmlRenderShape FollowingVisual(HtmlRenderPage page) => Assert.Single(
        page.Visuals.OfType<HtmlRenderShape>(),
        visual => visual.Source == "div#following" && visual.Shape.FillColor == FollowingColor);

    private static void AssertSvgPaint(OfficeRasterImage raster, bool stretch) {
        // The independent browser witness centers 64x32 paint inside the
        // taller concrete viewport; authored 'none' fills that viewport.
        // All samples stay several pixels away from antialiased paint edges.
        Assert.Equal(SvgColor, raster.GetPixel(32, 32));
        Assert.Equal(SvgColor, raster.GetPixel(32, 52));
        Assert.Equal(stretch ? SvgColor : BoxColor, raster.GetPixel(32, 20));
        Assert.Equal(stretch ? SvgColor : BoxColor, raster.GetPixel(32, 65));
    }
}
