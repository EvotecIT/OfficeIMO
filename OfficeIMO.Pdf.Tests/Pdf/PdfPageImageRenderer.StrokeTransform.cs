using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("9525 w q .00007874016 0 0 -.00007874016 55 125 cm 0 0 381000 762000 re S Q", 0.75D)]
    [InlineData("q 2 0 0 2 20 20 cm 3 w 0 0 10 10 re S Q", 6D)]
    public void RenderPage_ScalesOrdinaryStrokeWidthAtPaintTime(string content, double expectedWidth) {
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(BuildSingleStreamPdf(content));

        Assert.Equal(expectedWidth, Assert.Single(drawing.Shapes).Shape.StrokeWidth, 6);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
        Assert.Equal(OfficeColor.White, raster.GetPixel(5, 5));
    }
}
