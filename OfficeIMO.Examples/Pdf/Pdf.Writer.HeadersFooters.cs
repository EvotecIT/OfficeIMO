using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterHeadersFooters {
        public static void Example_Pdf_PageNumbers(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.WithPageNumbers.pdf");
            var options = new PdfCreateOptions {
                HeaderFormat = "OfficeIMO.Pdf headers and footers",
                HeaderAlign = PdfAlign.Left,
                HeaderOffsetY = 18,
                ShowHeader = true,
                FooterAlign = PdfAlign.Center,
                ShowPageNumbers = true
            };

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Headers and page numbers", PdfAlign.Left)
                .Paragraph(p => p.Text("The first page includes a header and the total page count."))
                .PanelParagraph(
                    p => p.Text("A page break starts a new section while retaining the document header and footer."),
                    new PdfPanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        PaddingX = 9,
                        PaddingY = 7
                    })
                .PageBreak()
                .H2("Continuation Page", PdfAlign.Left)
                .Paragraph(p => p.Text("Page numbers stay correct after an explicit page break."))
                .Paragraph(p => p.Text("Right-aligned continuation note."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))), options)
                .Meta(title: "OfficeIMO.Pdf Headers and Footers", author: "OfficeIMO")
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
