using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingCssColorTests {
    [Theory]
    [InlineData("rgb(255, 0, 128)", 255, 0, 128, 255)]
    [InlineData("rgb(100% 0% 50% / 25%)", 255, 0, 128, 64)]
    [InlineData("rgba(300, -20, 127.5, .5)", 255, 0, 128, 128)]
    [InlineData("hsl(120 100% 25%)", 0, 128, 0, 255)]
    [InlineData("hsla(0.5turn, 100%, 50%, 50%)", 0, 255, 255, 128)]
    [InlineData("hsl(-120deg 100% 50% / 0.75)", 0, 0, 255, 191)]
    public void OfficeColor_ParsesBoundedCssColorFunctions(string value, byte red, byte green, byte blue, byte alpha) {
        Assert.True(OfficeColor.TryParseCss(value, out OfficeColor color));
        Assert.Equal(OfficeColor.FromRgba(red, green, blue, alpha), color);
        Assert.Equal(color, OfficeColor.ParseCss(value));
    }

    [Theory]
    [InlineData("hwb(0 0% 0%)", 255, 0, 0, 255)]
    [InlineData("hwb(120deg 0% 0% / 50%)", 0, 255, 0, 128)]
    [InlineData("lab(100% 0 0)", 255, 255, 255, 255)]
    [InlineData("lab(0% 0 0 / 25%)", 0, 0, 0, 64)]
    [InlineData("lch(100% 0 270deg)", 255, 255, 255, 255)]
    [InlineData("oklab(1 0 0)", 255, 255, 255, 255)]
    [InlineData("oklch(0 0 180deg)", 0, 0, 0, 255)]
    [InlineData("color(srgb 1 0 .5 / .5)", 255, 0, 128, 128)]
    [InlineData("color(srgb-linear 1 0 0)", 255, 0, 0, 255)]
    [InlineData("color(display-p3 1 0 0)", 255, 0, 0, 255)]
    [InlineData("color(xyz-d65 0.95047 1 1.08883)", 255, 255, 255, 255)]
    [InlineData("color-mix(in srgb, red 25%, blue)", 64, 0, 191, 255)]
    [InlineData("color-mix(in srgb, rgb(255 0 0 / 50%), color(srgb 0 0 1))", 85, 0, 170, 192)]
    [InlineData("color-mix(in srgb, red 20%, blue 20%)", 128, 0, 128, 102)]
    public void OfficeColor_ParsesCssColorLevelFourFunctions(string value, byte red, byte green, byte blue, byte alpha) {
        Assert.True(OfficeColor.TryParseCss(value, out OfficeColor color));
        Assert.Equal(OfficeColor.FromRgba(red, green, blue, alpha), color);
    }

    [Theory]
    [InlineData("hsl(120 50 50)")]
    [InlineData("rgb(1 2)")]
    [InlineData("rgb(calc(1) 2 3)")]
    [InlineData("lab(50% 0)")]
    [InlineData("color(unknown 1 0 0)")]
    [InlineData("color-mix(in srgb, red, not-a-color)")]
    [InlineData("ff0000")]
    [InlineData("fff")]
    public void OfficeColor_RejectsUnsupportedOrMalformedCssColorFunctions(string value) {
        Assert.False(OfficeColor.TryParseCss(value, out _));
        Assert.Throws<FormatException>(() => OfficeColor.ParseCss(value));
    }

    [Fact]
    public void OfficeColor_CssHexRequiresTheCssHashPrefix() {
        Assert.True(OfficeColor.TryParseCss("#ff0000", out OfficeColor hex));
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), hex);
        Assert.True(OfficeColor.TryParseCss("red", out OfficeColor named));
        Assert.Equal(hex, named);
    }

    [Fact]
    public void SvgReader_UsesTheSharedCssColorParserForHslPaint() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'>"
            + "<rect width='20' height='10' fill='hsl(210 100% 40%)' stroke='hsl(30deg 100% 50% / 50%)'/>"
            + "</svg>";

        Assert.True(OfficeSvgDrawingReader.TryRead(
            System.Text.Encoding.UTF8.GetBytes(svg),
            out OfficeDrawing? drawing,
            out int unsupported));

        OfficeDrawingShape shape = Assert.Single(drawing!.Shapes);
        Assert.Equal(0, unsupported);
        Assert.Equal(OfficeColor.FromRgb(0, 102, 204), shape.Shape.FillColor);
        Assert.Equal(OfficeColor.FromRgba(255, 128, 0, 128), shape.Shape.StrokeColor);
    }
}
