using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class MarkdownPdfSvgInlineImageTests {
    [Fact]
    public void InlineSvgBadgeIsDrawnAsVectorsBesideItsText() {
        byte[] badge = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='60' height='20'>" +
            "<rect width='60' height='20' fill='#008000'/><text x='6' y='14' font-size='11' fill='#ffffff'>ok</text></svg>");
        MarkdownDocument document = MarkdownDocument.Parse("Build ![build ok](data:image/svg+xml;base64," +
            System.Convert.ToBase64String(badge) + ") today.");

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfCore.PdfDocument.Load(pdf).Reader.Images());
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Contains("today.", read.ExtractText());
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(read.Pages[0].ToDrawing(), background: OfficeColor.White);
        int green = 0;
        for (int y = 0; y < page.Height; y++)
            for (int x = 0; x < page.Width; x++)
                if (page.GetPixel(x, y) == OfficeColor.FromRgb(0, 128, 0)) green++;
        Assert.True(green > 100, "The badge was not drawn: " + green);
    }
}
