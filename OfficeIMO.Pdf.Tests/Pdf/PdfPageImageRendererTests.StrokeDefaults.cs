using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_StrokeDefaults_UsesPdfInitialMiterLimitForAcutePath() {
        byte[] pdf = BuildSingleStreamPdf("0 J 0 j 20 w 1 0 0 RG 80 20 m 100 120 l 120 20 l S");
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(100, 40));
        Assert.Equal(OfficeColor.White, raster.GetPixel(100, 20));
        Assert.Equal(10D, Assert.Single(drawing.Shapes).Shape.StrokeMiterLimit);
    }
}
