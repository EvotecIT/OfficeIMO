using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPrintAnnotationVisibilityTests {
    [Theory]
    [InlineData(4, true, true)]   // Print
    [InlineData(36, true, false)] // Print | NoView: printed, not shown on screen.
    [InlineData(6, false, false)] // Print | Hidden
    [InlineData(32, false, false)] // NoView without Print
    public void PrintedPagesUsePrintVisibilityForAnnotationFlags(int flags, bool printed, bool displayed) {
        PdfDocument document = PdfDocument.Load(AnnotatedPdf(flags));
        Assert.Equal(printed, IsRed(document.Render.PrintPage(1, new PdfPagePrintOptions { Dpi = 72 })));
        Assert.Equal(displayed, IsRed(document.Render.DisplayPage(1)));
    }

    private static bool IsRed(PdfPageRenderResult result) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(result.Bytes!, out OfficeRasterImage? image));
        int index = (image!.Height / 2 * image.Width + image.Width / 2) * 4;
        byte[] pixels = image.PixelBuffer;
        return pixels[index] > 200 && pixels[index + 1] < 60 && pixels[index + 2] < 60;
    }

    private static byte[] AnnotatedPdf(int flags) => Encoding.ASCII.GetBytes(string.Join("\n",
        "%PDF-1.4",
        "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
        "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
        "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Contents 4 0 R /Annots [5 0 R] >>", "endobj",
        "4 0 obj", "<< /Length 0 >>", "stream", "", "endstream", "endobj",
        "5 0 obj", "<< /Type /Annot /Subtype /Square /Rect [10 10 90 90] /F " + flags + " /AP << /N 6 0 R >> >>", "endobj",
        "6 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 80 80] /Length 23 >>", "stream", "1 0 0 rg 0 0 80 80 re f", "endstream", "endobj",
        "trailer", "<< /Root 1 0 R /Size 7 >>", "%%EOF"));
}
