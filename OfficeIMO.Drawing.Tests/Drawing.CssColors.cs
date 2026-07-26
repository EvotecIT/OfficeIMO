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
    [InlineData("hsl(120 50 50)")]
    [InlineData("rgb(1 2)")]
    [InlineData("rgb(calc(1) 2 3)")]
    [InlineData("lab(50% 0 0)")]
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
