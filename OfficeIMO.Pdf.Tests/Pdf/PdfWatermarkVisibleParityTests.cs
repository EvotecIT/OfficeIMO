using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(false, 5, 0)]
    [InlineData(false, 6, 0)]
    [InlineData(false, 5, 1)]
    [InlineData(false, 6, 1)]
    [InlineData(false, 5, 2)]
    [InlineData(false, 6, 2)]
    [InlineData(true, 5, 0)]
    [InlineData(true, 6, 0)]
    [InlineData(true, 5, 1)]
    [InlineData(true, 6, 1)]
    [InlineData(true, 5, 2)]
    [InlineData(true, 6, 2)]
    public void WatermarkVisibleParityPreservesSeparateFirstPageChoice(bool image, int start, int first) {
        byte[] png = PdfPngTestImages.CreateRgbPng(255, 0, 0);
        var document = PdfDocument.Create();
        document.Section(page => page.Content(content => content.Text("S1")));
        document.Section(page => {
            page.PageNumberStart(start);
            page.Options.UsePageNumberParityForHeaderFooter = true;
            if (image) {
                page.EvenPagesImageWatermark(new PdfImageWatermark(png, width: 22, height: 22));
                if (first == 1) page.FirstPageImageWatermark(new PdfImageWatermark(png, width: 31, height: 31));
                else if (first == 2) page.SuppressFirstPageImageWatermark();
            } else {
                page.EvenPagesWatermark("EVENMARK", font: PdfStandardFont.Courier);
                if (first == 1) page.FirstPageWatermark("FIRSTMARK", font: PdfStandardFont.TimesRoman);
                else if (first == 2) page.SuppressFirstPageTextWatermark();
            }
            page.Content(content => content.Text("S2A").PageBreak().Text("S2B"));
        });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 0; index < 2; index++) {
            bool even = (start + index) % 2 == 0;
            bool expected = index == 0 && first != 0 ? first == 1 : even;
            var page = pdf.GetPage(index + 2);
            if (image) {
                var images = page.GetImages().ToArray();
                Assert.Equal(expected ? 1 : 0, images.Length);
                if (expected) Assert.Equal(index == 0 && first == 1 ? 31D : 22D, images[0].BoundingBox.Width, 3);
            } else {
                string mark = index == 0 && first == 1 ? "FIRSTMARK" : "EVENMARK";
                Assert.Equal(expected, page.Text.Contains(mark));
                if (index == 0 && first == 1) Assert.DoesNotContain("EVENMARK", page.Text);
            }
        }
    }
}
