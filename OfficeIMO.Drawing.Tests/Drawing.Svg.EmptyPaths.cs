using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSvgEmptyPathsTests {
    [Theory]
    [InlineData("")]
    [InlineData(" d='' ")]
    [InlineData(" d='   ' ")]
    public void EmptyGlyphPathsDoNotReportUnsupportedPaint(string attribute) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40'>"
            + "<g role='img' aria-label=' '><path" + attribute + "/></g>"
            + "<rect x='5' y='5' width='20' height='10' fill='red'/></svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int unsupported));

        Assert.Equal(0, unsupported);
        Assert.NotNull(drawing);
        Assert.Single(drawing!.Elements);
    }

    [Theory]
    [InlineData("d='invalid'")]
    [InlineData("d='' filter='url(#paint)' ")]
    public void UnresolvedGeometryOrPaintProducingEffectsStillReportUnsupported(string attributes) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='100' height='40'>"
            + "<path " + attributes + "/></svg>");

        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out _, out int unsupported));

        Assert.True(unsupported > 0);
    }
}
