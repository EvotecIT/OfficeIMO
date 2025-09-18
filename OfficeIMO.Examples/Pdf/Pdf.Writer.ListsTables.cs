using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterListsTables {
        public static void Example_Pdf_BulletsAndTable(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.ListsTables.pdf");
            var rows = new[] {
                new [] { "Name", "Qty", "Price" },
                new [] { "Apples", "5", "$3.50" },
                new [] { "Bananas", "2", "$1.20" },
                new [] { "Cherries", "12", "$5.00" }
            };
            PdfDoc.Create()
                .H1("Simple Lists and Tables", PdfAlign.Center)
                .P("Below is a bullet list:")
                .Bullets(new[] { "First item", "Second item", "Third item" })
                .P(" ")
                .P("And a simple table (aligned columns):")
                .Table(rows)
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
