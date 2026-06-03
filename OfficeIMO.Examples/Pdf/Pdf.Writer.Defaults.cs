using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterDefaults {
        public static void Example_Pdf_DefaultStyles(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.DefaultStyles.pdf");
            var options = new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf default styles",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                DefaultTableStyle = TableStyles.Light()
            };

            var rows = new[] {
                new [] { "Metric", "Current", "Target" },
                new [] { "Runtime dependencies", "0", "0" },
                new [] { "Visual gates", "Growing", "Required for public claims" },
                new [] { "PowerShell wrapper", "PSWriteOffice", "Expose safe PDF operations" }
            };

            PdfDocument.Create(options)
                .Meta(title: "OfficeIMO.Pdf Default Styles", author: "OfficeIMO")
                .H1("Default Styles", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("This sample uses document-level defaults for text color, headers, footers, and the light table preset."))
                .PanelParagraph(
                    p => p.Text("The default table style should be good enough for a simple business report without every caller hand-tuning colors and padding."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        PaddingX = 9,
                        PaddingY = 7
                    })
                .Table(rows)
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
