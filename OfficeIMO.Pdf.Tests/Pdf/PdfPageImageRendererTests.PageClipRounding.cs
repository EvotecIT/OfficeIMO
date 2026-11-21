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

    [Fact]
    public void RenderPage_FitsPageClipWhoseIntersectionRoundsPastTheRightEdge() {
        // 0.072654 + (595.276 - 0.072654) exceeds 595.276 by one ulp, as an InDesign page-sized clip did.
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "0.072654 0 700 900 re W n\n1 0 0 rg\n100 100 600 50 re f",
            "<< >>",
            "/MediaBox [0 0 595.276 793.701]");

        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(300, 670));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(594, 670));
    }
}
