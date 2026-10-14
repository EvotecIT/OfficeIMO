using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    // Asymmetric 4x2 tile: moving the origin produces visibly different colors.
    private const string OverflowPatternTile = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAACCAIAAADwyuo0AAAAE0lEQVR4nGP4z8AARGAChv6DEQBwqQn3AyNm5wAAAABJRU5ErkJggg==";

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, 0)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, -30)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, 0)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, -30)]
    public void HtmlBackgroundOverflow_CropsPaintWithoutMovingTheTile(HtmlRenderIntentProfile profile, int left) {
        string html = OverflowPatternHtml($"position:relative;left:{left}px;");
        HtmlRenderResult result = RenderOverflowPattern(html, profile, HtmlRenderEncoder.Svg);
        HtmlRenderPage page = Assert.Single(result.Document.Pages);
        HtmlRenderImagePattern authored = Assert.Single(page.Visuals.OfType<HtmlRenderImagePattern>());
        OfficeDrawing drawing = page.CreateDrawing();
        OfficeImagePatternLayout cropped = Assert.Single(drawing.ImagePatterns).Layout;

        Assert.Equal(Math.Max(0D, 16D + left), cropped.Area.X);
        Assert.Equal(320D, cropped.Area.X + cropped.Area.Width, 6);
        Assert.Equal(authored.Pattern.Tile, cropped.Tile);
        Assert.Equal(authored.Pattern.RepeatX, cropped.RepeatX);
        Assert.Equal(authored.Pattern.RepeatY, cropped.RepeatY);
        Assert.Equal(authored.Pattern.HorizontalStep, cropped.HorizontalStep);
        Assert.Equal(authored.Pattern.VerticalStep, cropped.VerticalStep);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
        OfficeRasterImage viewport = OfficeDrawingRasterRenderer.Render(Assert.Single(RenderOverflowPattern(
            html, HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Png).Document.Pages).CreateDrawing(), 1D, OfficeColor.White);
        foreach (int x in new[] { 0, 16, 19, 23, 256, 319 }) {
            foreach (int y in new[] { 16, 17, 19, 20, 96 }) Assert.Equal(viewport.GetPixel(x, y), raster.GetPixel(x, y));
        }
        string svg = Encoding.UTF8.GetString(Assert.Single(result.ExportImages()).Bytes);
        Assert.Contains("<pattern", svg, StringComparison.Ordinal);
        Assert.Equal(320D, drawing.Width);
        Assert.Equal(230D, drawing.Height);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 },
            Assert.Single(RenderOverflowPattern(html, profile, HtmlRenderEncoder.Png).ExportImages()).Bytes.Take(8));
    }

    [Fact]
    public void HtmlBackgroundOverflow_OmitsPaintOutsideTheSurface() {
        HtmlRenderResult result = RenderOverflowPattern(OverflowPatternHtml("position:relative;left:400px;"),
            HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Svg);
        OfficeDrawing drawing = Assert.Single(result.Document.Pages).CreateDrawing();
        Assert.Empty(drawing.ImagePatterns);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.White, raster.GetPixel(319, 20));
        Assert.DoesNotContain("<pattern", Encoding.UTF8.GetString(Assert.Single(result.ExportImages()).Bytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HtmlBackgroundOverflow_RetainsLocalTransformAndClipPaint(bool transform) {
        string html = OverflowPatternHtml(transform ? "transform:translateX(-30px);" : "");
        if (!transform) html = html.Replace("<div></div>", "<section style='width:240px;height:80px;overflow:hidden'><div></div></section>");
        HtmlRenderResult result = RenderOverflowPattern(html, HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Svg);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(result.Document.Pages).CreateDrawing(), 1D, OfficeColor.White);
        Assert.NotEqual(OfficeColor.White, raster.GetPixel(20, 20));
        if (transform) Assert.NotEqual(OfficeColor.White, raster.GetPixel(256, 20));
        else Assert.Equal(OfficeColor.White, raster.GetPixel(256, 20));
        Assert.Single(result.ExportImages());
    }

    [Fact]
    public void HtmlBackgroundOverflow_PreservesVerticalFragmentPhase() {
        string html = OverflowPatternHtml("height:400px;");
        HtmlRenderResult result = RenderOverflowPattern(html, HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Svg);
        Assert.True(result.Document.Pages.Count > 1);
        Assert.Equal(result.Document.Pages.Count, result.ExportImages().Count);
        OfficeRasterImage firstPage = OfficeDrawingRasterRenderer.Render(result.Document.Pages[0].CreateDrawing(), 1D, OfficeColor.White);
        foreach (HtmlRenderPage page in result.Document.Pages) {
            HtmlRenderImagePattern authored = Assert.Single(EnumerateRenderVisuals(page.Scene).OfType<HtmlRenderImagePattern>());
            OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
            int first = (int)Math.Ceiling(Math.Max(16D, authored.Pattern.Area.Y));
            // Map back to the authored origin, rather than restarting the tile at each page top.
            for (int y = first; y < first + 4; y++) {
                int phase = ((int)(y - authored.Pattern.Tile.Y) % 4 + 4) % 4;
                Assert.Equal(firstPage.GetPixel(19, 17 + phase), raster.GetPixel(19, y));
            }
        }
    }

    [Theory]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, 0)]
    [InlineData(HtmlRenderIntentProfile.PrintPaged, -30)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, 0)]
    [InlineData(HtmlRenderIntentProfile.ScreenMediaPaged, -30)]
    public void HtmlImageOverflow_RetainsTheVisiblePartOfACroppedImage(HtmlRenderIntentProfile profile, int left) {
        string html = "<style>body{margin:0}</style><img src='data:image/png;base64," + OverflowPatternTile
            + $"' style='display:block;width:400px;height:80px;object-fit:cover;position:relative;left:{left}px'>";
        HtmlRenderResult result = RenderOverflowPattern(html, profile, HtmlRenderEncoder.Svg);
        HtmlRenderPage page = Assert.Single(result.Document.Pages);
        Assert.True(Assert.Single(page.Visuals.OfType<HtmlRenderImage>()).SourceCrop.HasCrop);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(page.CreateDrawing(), 1D, OfficeColor.White);
        OfficeRasterImage viewport = OfficeDrawingRasterRenderer.Render(Assert.Single(RenderOverflowPattern(
            html, HtmlRenderIntentProfile.ScreenViewport, HtmlRenderEncoder.Png).Document.Pages).CreateDrawing(), 1D, OfficeColor.White);
        foreach (int x in new[] { 0, 16, 19, 100, 256, 319 }) {
            foreach (int y in new[] { 16, 17, 40, 95, 96 }) Assert.Equal(viewport.GetPixel(x, y), raster.GetPixel(x, y));
        }
        Assert.Single(result.ExportImages());
        Assert.Single(RenderOverflowPattern(html, profile, HtmlRenderEncoder.Png).ExportImages());
    }

    [Fact]
    public void HtmlImageOverflow_OmitsAnImageOutsideTheSurface() {
        string html = "<style>body{margin:0}</style><img src='data:image/png;base64," + OverflowPatternTile
            + "' style='display:block;width:400px;height:80px;position:relative;left:400px'>";
        HtmlRenderResult result = RenderOverflowPattern(html, HtmlRenderIntentProfile.ScreenMediaPaged, HtmlRenderEncoder.Svg);
        Assert.DoesNotContain("<image", Encoding.UTF8.GetString(Assert.Single(result.ExportImages()).Bytes), StringComparison.Ordinal);
    }

    private static string OverflowPatternHtml(string extra) => "<style>body{margin:0}div{width:400px;height:80px;"
        + extra + "background-image:url(data:image/png;base64," + OverflowPatternTile
        + ");background-size:8px 4px;background-position:3px 1px;background-repeat:repeat}</style><div></div>";

    private static HtmlRenderResult RenderOverflowPattern(string html, HtmlRenderIntentProfile profile, HtmlRenderEncoder encoder) =>
        HtmlRenderEngine.Execute(HtmlConversionDocument.Parse(html), HtmlRenderRequest.Create(profile, encoder, new HtmlRenderOptions {
            PageSize = new OfficePageSize(320D / 96D, 230D / 96D),
            ViewportWidth = 320D,
            ViewportHeight = 230D,
            Margins = HtmlRenderMargins.All(16D),
            HonorCssPageRules = false
        }));
}
