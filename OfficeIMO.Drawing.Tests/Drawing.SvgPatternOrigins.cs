using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingSvgPatternOriginsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PatternContentStartsAtItsTileOrigin(bool stroke) {
        string shape = stroke ? "<path d='M0,20H100' stroke='url(#p)' stroke-width='10'/>" : "<rect width='100' height='40' fill='url(#p)'/>";
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40'><defs><pattern id='p' patternUnits='userSpaceOnUse' x='15' y='0' width='20' height='40'><rect width='10' height='40' fill='red'/><rect x='10' width='10' height='40' fill='blue'/></pattern></defs>" + shape + "</svg>";
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out var drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        var raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(17, 20));
        Assert.Equal(OfficeColor.Blue, raster.GetPixel(27, 20));
        if (stroke) {
            Assert.Equal(0, raster.GetPixel(17, 12).A);
            Assert.Equal(OfficeColor.Red, raster.GetPixel(17, 16));
            Assert.Equal(0, raster.GetPixel(17, 28).A);
        }
    }
}
