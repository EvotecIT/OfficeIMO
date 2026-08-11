using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("color-mix(in srgb, red -10%, blue)")]
    [InlineData("color-mix(in srgb, red 110%, blue)")]
    public void HtmlCssColors_RejectOutOfRangeColorMixPercentages(string value) {
        Assert.False(OfficeColor.TryParseCss(value, out _));
    }

    [Fact]
    public void HtmlRenderer_UsesSharedHslPaintAcrossTheSceneAndExporters() {
        const string html = "<div id='css-color' style='width:30px;height:14px;"
            + "background:hsl(210 100% 40%);border:2px solid hsl(30deg 100% 50% / 50%)'>Color</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            options);
        HtmlRenderShape fill = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#css-color" && item.Shape.FillColor.HasValue);
        HtmlRenderShape border = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#css-color" && item.Shape.StrokeColor.HasValue);
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);

        Assert.Equal(OfficeColor.FromRgb(0, 102, 204), fill.Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgba(255, 128, 0, 128), border.Shape.StrokeColor);
        Assert.Contains("#0066cc", svg, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(HtmlConversionDocument.Parse(html).ToPng(options));
    }

    [Fact]
    public void HtmlRenderer_UsesSharedCssColorLevelFourPaintAcrossTheSceneAndExporters() {
        const string html = "<div id='modern-color' style='width:30px;height:14px;"
            + "background:color-mix(in srgb, red 25%, blue);border:2px solid hwb(120 0% 0% / 50%)'>Color</div>";
        var options = new HtmlRenderOptions {
            ViewportWidth = 60D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        };

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            HtmlConversionDocument.Parse(html),
            options);
        HtmlRenderShape fill = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#modern-color" && item.Shape.FillColor.HasValue);
        HtmlRenderShape border = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#modern-color" && item.Shape.StrokeColor.HasValue);
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);

        Assert.Equal(OfficeColor.FromRgb(64, 0, 191), fill.Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgba(0, 255, 0, 128), border.Shape.StrokeColor);
        Assert.Contains("#4000bf", svg, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(HtmlConversionDocument.Parse(html).ToPng(options));
    }
}
