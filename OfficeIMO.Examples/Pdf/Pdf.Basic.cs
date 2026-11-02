using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class BasicPdf {
        public static void Example_Pdf_HelloWorld(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.pdf");
            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("OfficeIMO.Pdf — Hello World")
                .Paragraph(p => p.Text("This PDF uses the selected theme for fonts, spacing and page margins."))
                .Paragraph(p => p.Text("Text flows across the available page width, and the document embeds its resolved default font."))))
                .Meta(title: "Hello PDF", author: "OfficeIMO")
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
