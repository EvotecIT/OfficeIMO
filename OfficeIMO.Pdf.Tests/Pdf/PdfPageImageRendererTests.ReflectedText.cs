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
        byte[] pdf = BuildSingleStreamPdf(clip + $"BT /F1 10 Tf {spacing} Tc {matrix} 100 100 Tm (A) Tj ET",
            "<< /Font << /F1 5 0 R >> >>", font);
        OfficeDrawing actual = PdfPageImageRenderer.RenderPage(pdf);
        var expected = new OfficeDrawing(240, 200);
        var frame = new OfficeImageFrameTransform(horizontal ? -180 : 0, 100, 100, flipVertical: true);
        if (clipped) {
            expected.AddClippedPositionedText("A", 100, 90, 6, 12.5, 80, 180 - clipY,
                OfficeClipPath.Rectangle(40, 20), frame, new OfficeFontInfo("Courier New", 10),
                OfficeColor.Black, textAdvanceWidth: 6);
        } else {
            expected.AddPositionedText("A", 100, 90, 6, 12.5, frame,
                new OfficeFontInfo("Courier New", 10), OfficeColor.Black, textAdvanceWidth: 6);
        }
        byte[] pixels = OfficeDrawingRasterRenderer.Render(expected).GetPixels();
        Assert.Contains(pixels, channel => channel != 0);
        Assert.Equal(pixels, OfficeDrawingRasterRenderer.Render(actual).GetPixels());
        Assert.Contains("A", System.Text.Encoding.UTF8.GetString(actual.ExportImage(OfficeImageExportFormat.Svg).Bytes));
    }
}
