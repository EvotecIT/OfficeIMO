using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterHeadersFooters {
        public static void Example_Pdf_PageNumbers(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.WithPageNumbers.pdf");
            var options = new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf header/footer gate",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = 18,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true
            };

            PdfDoc.Create(options)
                .Meta(title: "OfficeIMO.Pdf Headers and Footers", author: "OfficeIMO")
                .H1("Header and Footer Baseline", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("Page one protects header placement, footer placement, and page number rendering on the first page."))
                .PanelParagraph(
                    p => p.Text("The same options should continue to render consistently after an explicit page break."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        PaddingX = 9,
                        PaddingY = 7
                    })
                .PageBreak()
                .H2("Continuation Page", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("Page two proves the raster harness compares more than the first page and guards the {page}/{pages} footer tokens."))
                .Paragraph(p => p.Text("Right-aligned continuation note."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
