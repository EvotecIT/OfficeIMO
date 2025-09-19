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
            var style = new PdfTableStyle {
                HeaderFill = PdfColor.LightGray,
                RowStripeFill = PdfColor.FromRgb(248, 248, 248),
                BorderColor = PdfColor.FromRgb(210, 210, 210),
                BorderWidth = 0.5
            };
            PdfDoc.Create()
                .H1("Simple Lists and Tables", PdfAlign.Center)
                .Paragraph(p => p.Text("Below is a bullet list:"))
                .Bullets(new[] { "First item", "Second item", "Third item" }, PdfAlign.Left, PdfColor.FromRgb(60, 60, 60))
                .Paragraph(p => p.Text(" "))
                .Paragraph(p => p.Text("And a simple table (aligned columns):"))
                .Table(rows, PdfAlign.Left, style)
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
