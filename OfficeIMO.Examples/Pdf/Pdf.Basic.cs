using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class BasicPdf {
        public static void Example_Pdf_HelloWorld(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "HelloWorld.OfficeIMO.Pdf.pdf");
            PdfDocument.Create()
                .Meta(title: "Hello PDF", author: "OfficeIMO")
                .H1("OfficeIMO.Pdf — Hello World")
                .Paragraph(p => p.Text("This PDF was generated with zero external dependencies using standard PDF fonts."))
                .Paragraph(p => p.Text("The layout uses simple vertical flow and Helvetica for readable default wrapping."))
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
