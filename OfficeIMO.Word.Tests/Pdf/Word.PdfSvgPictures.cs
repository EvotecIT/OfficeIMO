using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfSvgPictureTests {
    private const string Logo = "<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
        "<rect width='20' height='20' fill='#ff0000'/><rect x='20' width='20' height='20' fill='#0000ff'/></svg>";

    [Fact]
    public void SvgPicturesInBodyAndHeaderAreWrittenOnceAsVectors() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].PageSetup.Header.AddParagraph()
            .AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(40), OfficeLength.Points(20));
        for (int index = 0; index < 3; index++) {
            if (index > 0) document.AddPageBreak();
            document.AddParagraph("Page " + (index + 1));
            document.AddParagraph().AddImage(new MemoryStream(Encoding.UTF8.GetBytes(Logo)), "logo.svg", OfficeLength.Points(160), OfficeLength.Points(80));
        }

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(Regex.Matches(Encoding.GetEncoding("ISO-8859-1").GetString(pdf), @"/Subtype /Form /FormType 1 /BBox \[0 0 40 20\]"));
        PdfReadDocument read = PdfReadDocument.Open(pdf);
        Assert.Equal(3, read.Pages.Count);
        Assert.Contains("Page 3", read.ExtractText());
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(read.Pages[0].ToDrawing(), background: OfficeColor.White);
        int red = 0, blue = 0;
        for (int y = 0; y < page.Height; y++) {
            for (int x = 0; x < page.Width; x++) {
                OfficeColor color = page.GetPixel(x, y);
                if (color == OfficeColor.FromRgb(255, 0, 0)) red++;
                else if (color == OfficeColor.FromRgb(0, 0, 255)) blue++;
            }
        }
        // The body picture is 80 by 80 points per half, the header 20 by 20.
        Assert.InRange(red, 6000, 7200);
        Assert.InRange(blue, 6000, 7200);
    }
}
