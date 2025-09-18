using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterDefaults {
        public static void Example_Pdf_DefaultStyles(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.DefaultStyles.pdf");
            var options = new PdfOptions {
                DefaultTextColor = PdfColor.FromRgb(50, 50, 50),
                DefaultTableStyle = new PdfTableStyle {
                    HeaderFill = PdfColor.LightGray,
                    HeaderTextColor = PdfColor.Black,
                    RowStripeFill = PdfColor.FromRgb(245,245,245),
                    BorderColor = PdfColor.FromRgb(200,200,200),
                    BorderWidth = 0.5,
                    CellPaddingX = 6,
                    CellPaddingY = 2
                }
            };

            var rows = new[] {
                new [] { "Item", "Qty", "Cost" },
                new [] { "Pencils", "3", "$1.20" },
                new [] { "Notebooks", "2", "$4.00" },
                new [] { "Folders", "5", "$2.50" }
            };

            PdfDoc.Create(options)
                .H1("Defaults Demo", PdfAlign.Center, PdfColor.FromRgb(8,28,120))
                .P("Document uses default text color and table style.")
                .Table(rows) // picks up DefaultTableStyle
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}

