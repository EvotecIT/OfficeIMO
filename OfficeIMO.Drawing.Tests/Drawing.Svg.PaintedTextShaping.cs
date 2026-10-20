using System;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData("A", false)]
    [InlineData("A\u0301", true)]
    public void SvgReaderPaintedTextReportsIncompleteShapingWithoutDroppingInk(string text, bool expectedLoss) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='160' height='70'>"
            + "<defs><linearGradient id='ink'><stop stop-color='red'/><stop offset='1' stop-color='blue'/></linearGradient></defs>"
            + "<text x='8' y='40' font-family='Painted' font-size='22' fill='url(#ink)'>" + text + "</text></svg>";
        var options = new OfficeSvgDrawingReaderOptions();
        options.Fonts.Add("Painted", ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 0x0301));

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), options,
            out OfficeDrawing? drawing, out int unsupported));

        Assert.NotNull(drawing);
        Assert.Equal(expectedLoss, unsupported > 0);
        string exported = OfficeDrawingSvgExporter.ToSvg(drawing!);
        Assert.Contains("aria-label=\"" + text + "\"", exported, StringComparison.Ordinal);
        Assert.Contains("<linearGradient", exported, StringComparison.Ordinal);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing!);
        Assert.Contains(Enumerable.Range(0, raster.Height), y =>
            Enumerable.Range(0, raster.Width).Any(x => {
                OfficeColor pixel = raster.GetPixel(x, y);
                return pixel.A > 0 && (pixel.R != pixel.G || pixel.G != pixel.B);
            }));
    }
}
