using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData(false, "right 10px bottom 5px", 50, 31)]
    [InlineData(true, "right 10px bottom 5px", 50, 31)]
    [InlineData(false, "left 7px bottom", 7, 36)]
    [InlineData(true, "left 7px bottom", 7, 36)]
    [InlineData(false, "right calc(1em + 5px) bottom 0.5em", 45, 31)]
    [InlineData(true, "right calc(1em + 5px) bottom 0.5em", 45, 31)]
    [InlineData(false, "calc(50% - 10px) 25%", 20, 9)]
    [InlineData(true, "calc(50% - 10px) 25%", 20, 9)]
    [InlineData(false, "right 10% bottom 25%", 54, 27)]
    [InlineData(true, "right 10% bottom 25%", 54, 27)]
    public void HtmlBackgroundPositionEdges_UrlAndGradientTilesResolveAdmittedEdgeOffsets(
        bool gradient, string position, int expectedX, int expectedY) {
        string image = gradient ? "linear-gradient(red,red)"
            : "url(data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(255, 0, 0)) + ")";
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(BackgroundImageGeometryBox(
            "font-size:10px;background-image:" + image + ";background-size:20px 12px;"
            + "background-repeat:no-repeat;background-position:" + position));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        // Percentages, including edge offsets and math, use the free space
        // after the concrete 20x12 tile is removed from the 80x48 origin.
        Assert.Equal(OfficeColor.Red, raster.GetPixel(expectedX + 1, expectedY + 1));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(expectedX + 18, expectedY + 10));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX - 1, expectedY + 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 20, expectedY + 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 1, expectedY - 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 1, expectedY + 12));
        rendered.RequireNoLoss();
    }

    [Theory]
    [InlineData(false, false, false, "right 10% bottom 20%", 72, 40)]
    [InlineData(false, false, false, "right calc(10% + 2px) bottom calc(20% + 1px)", 70, 39)]
    [InlineData(false, true, false, "right 10% bottom 20%", 72, 40)]
    [InlineData(false, true, true, "right 10% bottom 20%", 72, 40)]
    [InlineData(true, false, false, "right 10% bottom 20%", 72, 40)]
    [InlineData(true, false, false, "right calc(10% + 2px) bottom calc(20% + 1px)", 70, 39)]
    [InlineData(true, true, true, "right 10% bottom 20%", 72, 40)]
    [InlineData(true, true, true, "right calc(10% + 2px) bottom calc(20% + 1px)", 70, 39)]
    public void HtmlBackgroundPositionEdges_ObjectPositionUsesTheSameFreeSpaceForEdgePercentages(
        bool svg, bool stylesheet, bool inlineBlock, string position, int expectedX, int expectedY) {
        // Keep real intrinsic image dimensions: object-fit:none paints this
        // 20x10 image in a 100x60 content box without scaling it.
        var scanlines = new byte[(20 * 3 + 1) * 10];
        for (int row = 0; row < 10; row++) {
            for (int column = 0; column < 20; column++) {
                scanlines[row * (20 * 3 + 1) + 1 + column * 3] = 255;
            }
        }
        string source = svg
            ? "data:image/svg+xml;base64," + Convert.ToBase64String(Encoding.UTF8.GetBytes(
                "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10' viewBox='0 0 20 10'>"
                + "<rect width='20' height='10' fill='red'/></svg>"))
            : "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreatePngWithScanlines(20, 10, 8, 2, scanlines));
        string declarations = "display:" + (inlineBlock ? "inline-block" : "block")
            + ";vertical-align:top;margin:0;width:100px;height:60px;object-fit:none;object-position:" + position;
        string image = "<img id='position-edge-object' class='object-reference' src='" + source + "'"
            + (stylesheet ? ">" : " style='" + declarations + "'>");
        HtmlRenderDocument rendered = RenderBackgroundImageGeometryCase(
            (stylesheet ? "<style>#position-edge-object{" + declarations + "}</style>" : "")
            + "<div style='font-size:0;line-height:0'>" + image + "</div>");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(Assert.Single(rendered.Pages).CreateDrawing());

        Assert.Equal(OfficeColor.Red, raster.GetPixel(expectedX + 1, expectedY + 1));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(expectedX + 18, expectedY + 8));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX - 1, expectedY + 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 20, expectedY + 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 1, expectedY - 1));
        Assert.Equal(OfficeColor.White, raster.GetPixel(expectedX + 1, expectedY + 10));
        rendered.RequireNoLoss();
    }
}
