using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgSymbolTests {
    [Fact]
    public void SymbolSafetyDoesNotUseRootDimensionsForNestedDefaults() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>"
            + "<svg width='20' height='20' viewBox='0 0 1000000 1000000'>"
            + "<defs><symbol id='tile'><rect width='1' height='1'/></symbol></defs><use href='#tile'/></svg></svg>";
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
    }

    [Fact]
    public void SymbolWithoutViewBoxKeepsUserCoordinatesAndClipsItsViewport() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='80' height='30'>"
            + "<defs><symbol id='tile'><rect width='12' height='10'/></symbol></defs>"
            + "<use href='#tile' x='2' y='2' width='6' height='8' fill='red'/>"
            + "<use href='#tile' x='20' y='2' width='24' height='12' fill='blue' preserveAspectRatio='none'/>"
            + "<use href='#tile' x='48' y='2' fill='lime'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
        var raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(4, 4));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(8, 4));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(4, 10));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(30, 4));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(34, 4));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(58, 4));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(62, 4));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SymbolPlacementUsesContainingViewportOrigin(bool symbolViewBox, bool nestedViewport) {
        string symbolAttributes = symbolViewBox ? "viewBox='0 0 4 4'" : "";
        string definition = $"<defs><symbol id='tile' {symbolAttributes}><rect width='4' height='4'/></symbol></defs>";
        string content = definition + "<use href='#tile' x='12' y='7' width='4' height='4' fill='red'/>";
        string svg = nestedViewport
            ? "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'><svg x='5' y='3' width='20' height='10' viewBox='10 5 20 10'>" + content + "</svg></svg>"
            : "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20' viewBox='10 5 40 20'>" + content + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        Assert.True(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
        var raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(nestedViewport ? 8 : 3, nestedViewport ? 6 : 3));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(nestedViewport ? 18 : 13, nestedViewport ? 11 : 8));
    }

    [Theory]
    [InlineData("viewBox=''", "width='12' height='10'")]
    [InlineData("viewBox='0 0 0 10'", "width='12' height='10'")]
    [InlineData("", "width='1000000' height='1000000'")]
    public void SymbolViewportRejectsMalformedViewBoxOrOversizedSurface(string symbolAttributes, string useAttributes) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='10'><defs>"
            + $"<symbol id='tile' {symbolAttributes}><rect width='12' height='10' fill='red'/></symbol></defs>"
            + $"<use href='#tile' {useAttributes}/><rect x='15' width='5' height='10' fill='lime'/></svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(1, unsupported);
        Assert.False(OfficeSvgDrawingReader.IsWithinSafetyLimits(Encoding.UTF8.GetBytes(svg)));
        var raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(4, 4));
        Assert.Equal(OfficeColor.Lime, raster.GetPixel(17, 4));
    }
}
