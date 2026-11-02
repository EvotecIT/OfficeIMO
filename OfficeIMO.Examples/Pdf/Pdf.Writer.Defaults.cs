using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterDefaults {
        public static void Example_Pdf_DefaultStyles(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.DefaultStyles.pdf");
            var options = new PdfCreateOptions {
                ShowPageNumbers = true
            };

            var rows = new[] {
                new [] { "Metric", "This quarter", "Target" },
                new [] { "Documents processed", "864", "900" },
                new [] { "Completed on time", "96%", "95%" },
                new [] { "Open follow-ups", "8", "0" }
            };

            PdfDocument.Create(pdf => pdf.Content(content => content
                .Title("Quarterly delivery review")
                .Paragraph(p => p.Text("This report uses the default Modern theme's type scale, page margins, table styling and list spacing."))
                .H2("Delivery results")
                .Table(rows)
                .H2("Next steps")
                .Bullets(new[] { "Complete the remaining follow-ups.", "Confirm the next quarter's delivery target." })), options)
                .Meta(title: "OfficeIMO.Pdf Default Styles", author: "OfficeIMO")
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
