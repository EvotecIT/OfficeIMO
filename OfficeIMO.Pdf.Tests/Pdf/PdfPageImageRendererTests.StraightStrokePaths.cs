using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RenderPage_RetainsStraightStrokesWithDegenerateBezierCorners(bool vertical, bool dashed) {
        string path = vertical
            ? "70 140 m 70 140 70 140 70 140 c 70 40 l 70 40 70 40 70 40 c"
            : "20 100 m 20 100 20 100 20 100 c 120 100 l 120 100 120 100 120 100 c";
        byte[] pdf = BuildSingleStreamPdf("q 22 20 96 116 re W n 1 0 0 RG 6 w "
            + (dashed ? "[8 12] 0 d " : string.Empty) + path + " S Q");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(vertical ? 70 : 25, vertical ? 65 : 100));
        Assert.Equal(dashed ? OfficeColor.White : OfficeColor.Red,
            raster.GetPixel(vertical ? 70 : 35, vertical ? 75 : 100));
        Assert.Equal(OfficeColor.White, raster.GetPixel(vertical ? 66 : 25, vertical ? 65 : 96));
        Assert.Equal(OfficeColor.White, raster.GetPixel(vertical ? 74 : 25, vertical ? 65 : 104));
        Assert.Equal(OfficeColor.White, raster.GetPixel(vertical ? 70 : 15, vertical ? 55 : 100));
        Assert.Equal(OfficeColor.White, raster.GetPixel(vertical ? 70 : 21, vertical ? 62 : 100));
    }
    [Theory]
    [InlineData(0.75D, 8D, 6, 25D)]
    [InlineData(2D, 2D, 4, -100D)]
    public void RenderPage_UniformTransformScalesTheStrokeBand(double scale, double width, int paintedWidth, double translateY) {
        string stream = FormattableString.Invariant(
            $"q {scale} 0 0 {scale} 0 {translateY} cm 1 0 0 RG {width} w 20 100 m 120 100 l S Q");
        byte[] pdf = BuildSingleStreamPdf(stream);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(pdf), background: OfficeColor.White);

        int top = 100 - paintedWidth / 2;
        Assert.Equal(OfficeColor.White, raster.GetPixel(60, top - 1));
        for (int y = top; y < top + paintedWidth; y++) Assert.Equal(OfficeColor.Red, raster.GetPixel(60, y));
        Assert.Equal(OfficeColor.White, raster.GetPixel(60, top + paintedWidth));
    }

    [Fact]
    public void RenderPage_UniformTransformScalesExactDashLengthsAndPhase() {
        byte[] pdf = BuildSingleStreamPdf(
            "q 0.75 0 0 0.75 0 25 cm 1 0 0 RG 8 w 0 J [8 12] 2 d 20 100 m 120 100 l S Q");
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(pdf), background: OfficeColor.White);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(17, 100));
        Assert.Equal(OfficeColor.White, raster.GetPixel(25, 100));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(31, 100));
    }

}
