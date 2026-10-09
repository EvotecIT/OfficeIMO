using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlBorders_AdjacentOpaqueFlexRowsRetainEveryBottomSeparator() {
        const string html = """
            <style>
              * { margin:0; padding:0; box-sizing:border-box }
              .row { display:flex; height:40px; border-bottom:1px solid #dee2e6 }
              .cell { height:39px; width:50%; background:#eef9fb }
              .cell + .cell { background:white }
            </style>
            <div class="row"><div class="cell"></div><div class="cell"></div></div>
            <div class="row"><div class="cell"></div><div class="cell"></div></div>
            <div class="row"><div class="cell"></div><div class="cell"></div></div>
            """;

        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 125D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));

        OfficeColor separator = OfficeColor.FromRgb(0xDE, 0xE2, 0xE6);
        foreach (int y in new[] { 39, 79, 119 }) {
            Assert.Equal(separator, raster!.GetPixel(15, y));
            Assert.Equal(separator, raster.GetPixel(90, y));
        }
        Assert.Equal(OfficeColor.FromRgb(0xEE, 0xF9, 0xFB), raster!.GetPixel(15, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(90, 40));
    }

    [Theory]
    [InlineData("solid")]
    [InlineData("dashed")]
    [InlineData("dotted")]
    [InlineData("double")]
    [InlineData("inset")]
    [InlineData("outset")]
    [InlineData("groove")]
    [InlineData("ridge")]
    public void HtmlBorders_SingleEdgePaintUsesItsFullBandBeforeAnOpaqueSibling(string borderStyle) {
        string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style>"
            + "<div style='width:120px;height:40px;border-bottom:6px " + borderStyle + " #808080;background:white'></div>"
            + "<div style='width:120px;height:12px;background:#00ff00'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 120D,
            ViewportHeight = 56D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));

        for (int y = 34; y < 40; y++) {
            int painted = Enumerable.Range(8, 104).Count(x => raster!.GetPixel(x, y) != OfficeColor.White);
            if (borderStyle == "double" && y is 36 or 37) Assert.Equal(0, painted);
            else Assert.True(painted > 0, borderStyle + " did not paint its border-box band at y=" + y);
        }
        Assert.Equal(OfficeColor.White, raster!.GetPixel(60, 33));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(60, 40));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(60, 51));
    }

    [Theory]
    [InlineData("solid", false)]
    [InlineData("dashed", false)]
    [InlineData("dotted", false)]
    [InlineData("double", false)]
    [InlineData("inset", false)]
    [InlineData("outset", false)]
    [InlineData("groove", false)]
    [InlineData("ridge", false)]
    [InlineData("solid", true)]
    [InlineData("double", true)]
    [InlineData("groove", true)]
    public void HtmlBorders_RoundedPaintStaysWithinUniformAndAsymmetricBorderBoxes(string borderStyle, bool asymmetric) {
        string widths = asymmetric ? "4px 8px 10px 6px" : "6px";
        string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style>"
            + "<div style='margin:8px;width:64px;height:44px;border-width:" + widths
            + ";border-style:" + borderStyle + ";border-color:#808080;border-radius:12px 8px / 10px 16px;background:white'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));

        Assert.Contains(Enumerable.Range(8, 64).SelectMany(x => Enumerable.Range(8, 44).Select(y => raster!.GetPixel(x, y))),
            pixel => pixel != OfficeColor.White);
        for (int y = 0; y < 60; y++) {
            for (int x = 0; x < 80; x++) {
                if (x < 8 || x >= 72 || y < 8 || y >= 52) Assert.Equal(OfficeColor.White, raster!.GetPixel(x, y));
            }
        }
        Assert.Equal(OfficeColor.White, raster!.GetPixel(8, 8));
        Assert.Equal(OfficeColor.White, raster.GetPixel(40, 30));
    }
}
