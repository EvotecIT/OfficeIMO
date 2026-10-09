using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("block", "width:100px;height:0;border-bottom:1px solid red", 100D, 1D, 20, 8)]
    [InlineData("flex", "width:100px;height:0;border-bottom:1px solid red", 100D, 1D, 20, 8)]
    [InlineData("grid", "width:100px;height:0;border-bottom:1px solid red", 100D, 1D, 20, 8)]
    [InlineData("inline-block", "width:100px;height:0;border-bottom:1px solid red", 100D, 1D, 20, 8)]
    [InlineData("block", "width:100px;height:1px;border:1px solid red", 100D, 2D, 20, 8)]
    [InlineData("flex", "width:100px;height:1px;border:1px solid red", 100D, 2D, 20, 8)]
    [InlineData("grid", "width:100px;height:1px;border:1px solid red", 100D, 2D, 20, 8)]
    [InlineData("inline-block", "width:100px;height:1px;border:1px solid red", 100D, 2D, 20, 8)]
    [InlineData("block", "width:0;height:40px;border-left:4px solid red", 4D, 40D, 8, 20)]
    [InlineData("flex", "width:0;height:40px;border-left:4px solid red", 4D, 40D, 8, 20)]
    [InlineData("grid", "width:0;height:40px;border-left:4px solid red", 4D, 40D, 8, 20)]
    [InlineData("inline-block", "width:0;height:40px;border-left:4px solid red", 4D, 40D, 8, 20)]
    public void HtmlBorders_EmptyBorderBoxesRetainTheirBorderUsedSize(string layout, string css, double width, double height, int pixelX, int pixelY) {
        string display = layout == "inline-block" ? "display:inline-block;vertical-align:top;" : string.Empty;
        string parent = layout == "flex" ? "display:flex;align-items:flex-start;"
            : layout == "grid" ? "display:grid;grid-template-columns:120px;align-items:start;justify-items:start;"
            : "font-size:0;line-height:0;";
        string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style><div style='" + parent + "'>"
            + "<div id='divider' style='margin:8px;background:white;" + display + css + "'></div></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 140D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.White
        };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape background = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#divider" && shape.Shape.FillColor == OfficeColor.White);
        Assert.Equal(width, background.Width, 3);
        Assert.Equal(height, background.Height, 3);
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(pixelX, pixelY));
    }

    [Fact]
    public void HtmlBorders_MaximumBorderBoxSizesCannotRemoveBorderAndPaddingInsets() {
        const string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style>"
            + "<div id='minimum' style='width:0;height:0;max-width:0;max-height:0;padding:3px;border:2px solid red;background:white'></div>";
        var options = new HtmlRenderOptions { ViewportWidth = 20D, ViewportHeight = 20D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape background = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#minimum" && shape.Shape.FillColor == OfficeColor.White);
        Assert.Equal(10D, background.Width, 3);
        Assert.Equal(10D, background.Height, 3);
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(5, 0));
        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
    }

    [Fact]
    public void HtmlBorders_ZeroSizeControlsPreserveBorderAndPaddingInsets() {
        const string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style>"
            + "<input id='control' type='text' style='margin:8px;width:0;height:0;max-width:0;max-height:0;padding:3px;border:2px solid red;background:white'>";
        var options = new HtmlRenderOptions { ViewportWidth = 40D, ViewportHeight = 40D, Margins = HtmlRenderMargins.All(0D) };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderFormField field = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Scene).OfType<HtmlRenderFormField>(),
            visual => visual.Source == "input#control");
        Assert.Equal(10D, field.Width, 3);
        Assert.Equal(10D, field.Height, 3);
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(13, 8));
        Assert.Equal(OfficeColor.White, raster.GetPixel(13, 13));
    }

    [Theory]
    [InlineData("2px")]
    [InlineData("12px / 2px")]
    public void HtmlBorders_ThickUniformStrokeRetainsItsOuterRadiusClip(string radius) {
        string html = "<style>*{margin:0;padding:0;box-sizing:border-box}</style>"
            + "<div id='rounded-band' style='margin:8px;width:60px;height:40px;background:white;border:10px solid #808080;border-radius:" + radius + "'></div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 80D,
            ViewportHeight = 60D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.FromRgb(0xE0, 0xF0, 0xFF)
        };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderShape border = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#rounded-band" && shape.Shape.StrokeColor.HasValue);
        Assert.NotNull(border.Shape.ClipPath);
        Assert.Equal(60D, border.Shape.ClipPath!.Width, 3);
        Assert.Equal(40D, border.Shape.ClipPath.Height, 3);
        Assert.True(OfficePngReader.TryDecode(HtmlConversionDocument.Parse(html).ToPng(options), out OfficeRasterImage? raster));
        Assert.NotEqual(OfficeColor.FromRgb(0x80, 0x80, 0x80), raster!.GetPixel(8, 8));
        Assert.Equal(OfficeColor.FromRgb(0x80, 0x80, 0x80), raster.GetPixel(30, 8));
    }
}
