using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterListsTables {
        public static void Example_Pdf_BulletsAndTable(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.ListsTables.pdf");
            var rows = new[] {
                new [] { "Item", "Qty", "Unit", "Total", "Notes" },
                new [] { "Monitoring seats", "25", "$4.50", "$112.50", "Monthly report-ready subscription line." },
                new [] { "Security review", "1", "$250.00", "$250.00", "Includes a short executive summary and remediation list." },
                new [] { "Documentation pack", "3", "$35.00", "$105.00", "Generated attachments for PSWriteOffice hand-off." },
                new [] { "Total", "", "", "$467.50", "Ready for approval." }
            };

            var style = new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(32, 76, 120),
                HeaderTextColor = PdfColor.White,
                TextColor = PdfColor.FromRgb(31, 41, 55),
                FooterFill = PdfColor.FromRgb(232, 241, 248),
                FooterTextColor = PdfColor.FromRgb(25, 55, 85),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                BorderColor = PdfColor.FromRgb(210, 218, 226),
                BorderWidth = 0.5,
                CellPaddingX = 6,
                CellPaddingY = 5,
                HeaderRowCount = 1,
                FooterRowCount = 1,
                Caption = "Table 1. Example line items",
                CaptionColor = PdfColor.FromRgb(80, 90, 100),
                CaptionFontSize = 8.5,
                CaptionSpacingAfter = 5,
                SpacingBefore = 6,
                SpacingAfter = 14,
                RightAlignNumeric = true,
                Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Left
                },
                ColumnWidthPoints = new List<double?> { 110, 38, 58, 68, 170 },
                AutoFitColumns = false
            };

            PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf lists and tables",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Lists and Tables", author: "OfficeIMO")
                .H1("Lists and Tables", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("A compact report sample for list rhythm, numeric alignment, footer rows, and wrapped table notes."))
                .Bullets(new[] {
                    "Report-friendly list spacing",
                    "Right-aligned quantities and amounts",
                    "Footer row that remains visually distinct"
                }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
                .Numbered(new[] {
                    "Collect line items.",
                    "Render the table with stable column widths.",
                    "Review the generated PDF through the raster gate."
                }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
                .Table(rows, PdfAlign.Left, style)
                .Paragraph(p => p.Text("End of lists and tables sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
