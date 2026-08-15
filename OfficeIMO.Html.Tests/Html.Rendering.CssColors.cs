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

    [Fact]
    public void HtmlRenderer_DiagnosesUnsupportedForegroundAndBackgroundColorValues() {
        const string html = "<div style='color:color(from red srgb r g b);background-color:color(from blue srgb r g b)'>Fallback</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);

        Assert.Equal(2, rendered.Diagnostics.Count(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ColorValueUnsupported));
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Detail == "color=color(from red srgb r g b)");
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Detail == "background-color=color(from blue srgb r g b)");
        Assert.Throws<HtmlConversionException>(() => HtmlRenderTestDriver.Render(
            html,
            new HtmlRenderOptions { FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss }));
        Assert.Contains(HtmlRenderDiagnosticCodes.ColorValueUnsupported, HtmlRenderDiagnosticCodes.All);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.ColorValueUnsupported, out _));
    }

    [Theory]
    [InlineData("red url('{0}') no-repeat")]
    [InlineData("url('{0}') no-repeat red")]
    public void HtmlRenderer_PreservesColorInCombinedBackgroundShorthand(string shorthand) {
        const string pixelPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII=";
        string dataUri = "data:image/png;base64," + pixelPng;
        string html = "<div id='combined' style=\"width:20px;height:12px;background:"
            + string.Format(System.Globalization.CultureInfo.InvariantCulture, shorthand, dataUri)
            + "\">Color</div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss,
            BackgroundColor = OfficeColor.Transparent
        });
        HtmlRenderShape fill = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#combined" && item.Shape.FillColor.HasValue);

        Assert.Equal(OfficeColor.Red, fill.Shape.FillColor);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.ColorValueUnsupported);
    }

    [Theory]
    [InlineData("lab(50 200 0)", "lab(50 125 0)")]
    [InlineData("lab(50 -200 0)", "lab(50 -125 0)")]
    [InlineData("lch(50 250 300)", "lch(50 150 300)")]
    [InlineData("oklab(.5 .8 0)", "oklab(.5 .4 0)")]
    [InlineData("oklch(.5 .8 300)", "oklch(.5 .4 300)")]
    public void HtmlCssColors_NumericLabAxesAndPositiveChromaAreNotClampedToPercentageRanges(string extended, string percentageBoundary) {
        Assert.True(OfficeColor.TryParseCss(extended, out OfficeColor extendedColor));
        Assert.True(OfficeColor.TryParseCss(percentageBoundary, out OfficeColor boundaryColor));

        Assert.NotEqual(boundaryColor, extendedColor);
    }
}
