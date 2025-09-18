using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterListsTables {
        public static void Example_Pdf_BulletsAndTable(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.ListsTables.pdf");
            PdfDoc.Create()
                .H1("Simple Lists and Tables", PdfAlign.Center)
                .P("Below is a bullet list:")
                .Bullets(new[] { "First item", "Second item", "Third item" })
                .P("And a simple table (monospaced grid):")
                .P(" ")
                .P("Name | Qty | Price")
                .P("---- | --- | -----")
                .P("Apples | 5 | $3.50")
                .P("Bananas | 2 | $1.20")
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}

