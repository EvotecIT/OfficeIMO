using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData(true, false, 0)]
    [InlineData(true, true, 0)]
    [InlineData(false, false, 0)]
    [InlineData(false, true, 0)]
    [InlineData(true, true, 1)]
    [InlineData(false, true, 1)]
    public void RenderPage_PreservesReflectedTextPaintAndClip(bool horizontal, bool clipped, int spacing) {
        const string font = "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj";
        string matrix = horizontal ? "-1 0 0 1" : "1 0 0 -1";
        int clipY = horizontal ? 100 : 80;
        string clip = clipped ? $"80 {clipY} 40 20 re W n " : "";
        byte[] pdf = BuildSingleStreamPdf(clip + $"BT /F1 10 Tf {spacing} Tc {matrix} 100 100 Tm (AB) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(pdf);
        var expected = new OfficeDrawing(240, 200);
        for (int index = 0; index < (spacing == 0 ? 1 : 2); index++) {
            double x = 100 + (horizontal ? -7 : 7) * index;
            var frame = new OfficeImageFrameTransform(horizontal ? -180 : 0, x, 100, flipVertical: true);
            string text = spacing == 0 ? "AB" : index == 0 ? "A" : "B";
            double width = spacing == 0 ? 12 : 6;
            if (clipped) {
                expected.AddClippedPositionedText(text, x, 90, width, 12.5, 80, 180 - clipY,
                    OfficeClipPath.Rectangle(40, 20), frame, new OfficeFontInfo("Courier New", 10),
                    OfficeColor.Black, textAdvanceWidth: width);
            } else {
                expected.AddPositionedText(text, x, 90, width, 12.5, frame,
                    new OfficeFontInfo("Courier New", 10), OfficeColor.Black, textAdvanceWidth: width);
            }
        }
        byte[] pixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(pixels, OfficeDrawingRasterRenderer.Render(actual).GetPixels());
        Assert.Contains("A", System.Text.Encoding.UTF8.GetString(actual.ExportImage(OfficeImageExportFormat.Svg).Bytes));
    }
}
