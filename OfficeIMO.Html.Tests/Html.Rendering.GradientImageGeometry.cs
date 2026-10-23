using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("linear-gradient(to bottom right,red,blue)")]
    [InlineData("repeating-linear-gradient(to right,red 0px 4px,blue 4px 8px)")]
    [InlineData("radial-gradient(ellipse closest-side at center,red,blue)")]
    [InlineData("repeating-radial-gradient(circle at center,red 0px 4px,blue 4px 8px)")]
    [InlineData("conic-gradient(from 90deg,red,blue,red)")]
    [InlineData("repeating-conic-gradient(red 0deg 45deg,blue 45deg 90deg)")]
    public void HtmlGradientImageGeometry_SixFamiliesUseTheConcretePositionedTile(string image) {
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "background-image:" + image + ";background-size:20px 12px;"
            + "background-position:25% 50%;background-repeat:no-repeat"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        // Position percentages apply to the space remaining after tile sizing.
        Assert.NotEqual(OfficeColor.White, raster.GetPixel(16, 19));
        Assert.NotEqual(OfficeColor.White, raster.GetPixel(33, 28));
        Assert.Equal(OfficeColor.White, raster.GetPixel(14, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(16, 17));
        Assert.Equal(OfficeColor.White, raster.GetPixel(16, 30));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("20px auto", 20, 48)]
    [InlineData("auto 12px", 80, 12)]
    [InlineData("contain", 80, 48)]
    [InlineData("cover", 80, 48)]
    public void HtmlGradientImageGeometry_AutoAxesDoNotInventAnIntrinsicAspectRatio(
        string size, int expectedWidth, int expectedHeight) {
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "background-image:linear-gradient(red,red);background-size:" + size
            + ";background-repeat:no-repeat"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(expectedWidth - 2, expectedHeight - 2));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedWidth + 1, expectedHeight / 2));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedWidth / 2, expectedHeight + 1));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlGradientImageGeometry_RoundDoesNotRescaleTheOtherAutoAxis() {
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "background-image:linear-gradient(to bottom,red 0% 50%,blue 50% 100%);"
            + "background-size:30px auto;background-repeat:round no-repeat"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        // Rounding three tiles to 80/3 pixels must keep the gradient's 48px height.
        Assert.Equal(OfficeColor.Red, raster.GetPixel(4, 22));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(4, 26));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(75, 45));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlGradientImageGeometry_SpaceUsesTheOriginInsteadOfTheLargerClip() {
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "height:20px;padding:10px;border:2px solid green;"
            + "background-image:linear-gradient(red,red);background-size:18px 14px;"
            + "background-origin:content-box;background-clip:border-box;"
            + "background-repeat:space;background-position:right bottom"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        // Four tiles share the 80px content origin. The 20px vertical origin fits
        // one tile, so its bottom position remains effective on that axis.
        Assert.Equal(OfficeColor.Red, raster.GetPixel(13, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(31, 19));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 19));
        Assert.Equal(OfficeColor.White, raster.GetPixel(13, 15));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData("0px 8px")]
    [InlineData("8px 0px")]
    public void HtmlGradientImageGeometry_ZeroTileAxisSuppressesPaintWithoutSpendingTheBudget(string size) {
        string empty = BackgroundImageGeometryBox("background-image:linear-gradient(red,red);background-size:" + size);
        string visible = BackgroundImageGeometryBox(
            "background-image:linear-gradient(blue,blue);background-size:20px 12px");
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(empty + visible, maximumTiles: 16);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(75, 90));
        rendered.RequireNoLoss();
    }

    [Fact]
    public void HtmlGradientImageGeometry_GradientAndUrlTilesShareTheOperationBudgetAndStrictLoss() {
        string blue = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(0, 0, 255));
        string first = BackgroundImageGeometryBox(
            "width:40px;height:20px;background-image:linear-gradient(red,red);background-size:10px 10px");
        string second = BackgroundImageGeometryBox(
            "width:40px;height:20px;background-image:url(data:image/png;base64," + blue + ");background-size:10px 10px");
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(first + second, maximumTiles: 10);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(35, 15));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(5, 25));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 35));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.BackgroundImageTileLimitExceeded);
        Assert.ThrowsAny<InvalidOperationException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlGradientImageGeometry_TileBoundsKeepLayerOrderAndAlphaComposition() {
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "background-image:linear-gradient(rgba(255,0,0,0.5),rgba(255,0,0,0.5)),linear-gradient(blue,blue);"
            + "background-size:20px 12px,40px 24px;background-repeat:no-repeat"));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());
        OfficeColor overlap = raster.GetPixel(10, 6);

        Assert.InRange((int)overlap.R, 127, 129);
        Assert.Equal(0, overlap.G);
        Assert.InRange((int)overlap.B, 126, 128);
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(30, 6));
        Assert.Equal(OfficeColor.White, raster.GetPixel(60, 6));
        rendered.RequireNoLoss();
    }

    private static string BackgroundImageGeometryBox(string declarations) =>
        "<div style='width:80px;height:48px;" + declarations + "'></div>";

    private static HtmlRenderDocument RenderBackgroundImageGeometryCase(string content, int maximumTiles = 16384) =>
        HtmlRenderTestDriver.Render("<style>html,body{margin:0}</style>" + content, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Continuous,
            ViewportWidth = 128D,
            ViewportHeight = 128D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White,
            MaxBackgroundImageTiles = maximumTiles
        });
}
