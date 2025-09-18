using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterHeadersFooters {
        public static void Example_Pdf_PageNumbers(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.WithPageNumbers.pdf");
            var options = new PdfOptions { ShowPageNumbers = true, FooterAlign = PdfAlign.Center, FooterFormat = "Page {page}/{pages}", DefaultFont = PdfStandardFont.Courier };
            PdfDoc.Create(options)
                .H1("Report", PdfAlign.Center)
                .P("This demonstrates page numbers rendered in the footer.")
                .PageBreak()
                .H2("Second Page", PdfAlign.Right)
                .P("Right-aligned paragraph on page 2.", PdfAlign.Right)
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}

