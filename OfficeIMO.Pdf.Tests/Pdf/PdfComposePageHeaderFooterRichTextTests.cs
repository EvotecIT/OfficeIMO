using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf {
    public partial class PdfComposePageOptionsTests {
        [Fact]
        public void HeaderFooterRichText_RendersStylesAndStyledPageTokens() {
            var options = new PdfOptions {
                CompressContentStreams = false,
                HeaderFont = PdfStandardFont.Helvetica,
                FooterFont = PdfStandardFont.Helvetica
            };
            var doc = PdfDocument.Create(options)
                .Header(header => header.Text(text => text
                    .Run(TextRun.Bolded("Rich header ", PdfColor.FromRgb(255, 0, 0), fontSize: 14, backgroundColor: PdfColor.FromRgb(255, 255, 0)))
                    .CurrentPage(TextRun.Italicized(string.Empty, PdfColor.FromRgb(0, 0, 255), fontSize: 11))
                    .Text("/")
                    .TotalPages(TextRun.Underlined(string.Empty, PdfColor.FromRgb(0, 128, 0), fontSize: 11))))
                .Footer(footer => footer.Text(text => text
                    .Run(TextRun.Strikethrough("Rich footer ", PdfColor.FromRgb(128, 0, 128), fontSize: 9))
                    .CurrentPage(TextRun.Superscript(string.Empty, fontSize: 9))))
                .Paragraph(p => p.Text("Body content."));

            byte[] bytes = doc.ToBytes();
            string rawPdf = Encoding.ASCII.GetString(bytes);
            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            var page = pdf.GetPage(1);
            string text = Normalize(page.Text);

            Assert.Contains("Richheader1/1", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Richfooter1", text, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(page.Letters, letter => letter.Value == "R" && letter.FontName != null && letter.FontName.Contains("Bold", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(page.Letters, letter => letter.Value == "1" && letter.FontName != null && letter.FontName.Contains("Oblique", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(page.Letters, letter => letter.Value == "R" && letter.PointSize > 13D);
            Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
            Assert.Contains("0 0 1 rg", rawPdf, StringComparison.Ordinal);
            Assert.Contains("1 1 0 rg", rawPdf, StringComparison.Ordinal);
            Assert.Contains("0 0.502 0 RG", rawPdf, StringComparison.Ordinal);
            Assert.Contains("0.502 0 0.502 RG", rawPdf, StringComparison.Ordinal);
            Assert.Contains(" Ts", rawPdf, StringComparison.Ordinal);
        }

        [Fact]
        public void HeaderFooterRichText_UsesFirstEvenAndDefaultVariants() {
            byte[] bytes = PdfDocument.Create()
                .Header(header => header
                    .Text(text => text.Run(TextRun.Bolded("Odd rich")))
                    .FirstPageText(text => text.Run(TextRun.Italicized("First rich")))
                    .EvenPagesText(text => text.Run(TextRun.Underlined("Even rich"))))
                .Footer(footer => footer
                    .Text(text => text.Run(TextRun.Bolded("Odd footer")))
                    .FirstPageText(text => text.Run(TextRun.Italicized("First footer")))
                    .EvenPagesText(text => text.Run(TextRun.Underlined("Even footer"))))
                .Paragraph(p => p.Text("Page one"))
                .PageBreak()
                .Paragraph(p => p.Text("Page two"))
                .PageBreak()
                .Paragraph(p => p.Text("Page three"))
                .ToBytes();

            using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
            Assert.Contains("Firstrich", Normalize(pdf.GetPage(1).Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Firstfooter", Normalize(pdf.GetPage(1).Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Evenrich", Normalize(pdf.GetPage(2).Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Evenfooter", Normalize(pdf.GetPage(2).Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddrich", Normalize(pdf.GetPage(3).Text), StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Oddfooter", Normalize(pdf.GetPage(3).Text), StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HeaderFooterRichText_RejectsInteractiveAndInlineRuns() {
            var link = TextRun.Link("Link", "https://example.com");
            var inline = TextRun.Inline(new PdfInlineBox(12, 8));

            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Header(header => header.Text(text => text.Run(link))));
            Assert.Throws<ArgumentException>(() =>
                PdfDocument.Create().Footer(footer => footer.Text(text => text.Run(inline))));
            Assert.Throws<ArgumentException>(() =>
                FooterSegment.PageNumber(TextRun.Tab()));
        }
    }
}
