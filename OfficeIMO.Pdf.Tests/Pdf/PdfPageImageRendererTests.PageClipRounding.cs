using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("-0.0005 0 240.0745 200.074 re f")]
    [InlineData("0 0 240.0745 200.074 re f")]
    [InlineData("0 0 240.074 200.0745 re f")]
    [InlineData("0 -0.0005 240.074 200.0745 re f")]
    public void RenderPage_RetainsCanvasFillWhenFractionalGeometrySlightlyCrossesPageEdges(string rectangle) {
        // Independent PDF content reproduces the small geometry/MediaBox rounding
        // differences produced when a fractional HTML surface is serialized.
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "1 0 0 rg\n" + rectangle,
            "<< >>",
            "/MediaBox [0 0 240.074 200.074]");

        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(1, 1));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(120, 100));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(239, 199));
    }
}
