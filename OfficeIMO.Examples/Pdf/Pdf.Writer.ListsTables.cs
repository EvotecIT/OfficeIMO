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
                new [] { "Documentation pack", "3", "$35.00", "$105.00", "Generated attachments for the monthly report." },
                new [] { "Total", "", "", "$467.50", "Ready for approval." }
            };

            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = 6;
            style.CellPaddingY = 5;
            style.CellPaddingTop = 5;
            style.CellPaddingBottom = 5;
            style.HeaderRowCount = 1;
            style.FooterRowCount = 1;
            style.Caption = "Table 1. Example line items";
            style.CaptionFontSize = 8.5;
            style.CaptionSpacingAfter = 5;
            style.SpacingBefore = 6;
            style.SpacingAfter = 14;
            style.RightAlignNumeric = true;
            style.Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Right,
                    PdfColumnAlign.Left
                };
            style.ColumnWidthWeights = new List<double> { 110, 38, 58, 68, 170 };
            style.AutoFitColumns = false;

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Lists and Tables", PdfAlign.Left)
                .Paragraph(p => p.Text("A compact report sample for list rhythm, numeric alignment, footer rows, and wrapped table notes."))
                .Bullets(new[] {
                    "Report-friendly list spacing",
                    "Right-aligned quantities and amounts",
                    "Footer row that remains visually distinct"
                }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
                .Numbered(new[] {
                    "Collect line items.",
                    "Render the table with stable column widths.",
                    "Review the saved PDF before sharing it."
                }, PdfAlign.Left, PdfColor.FromRgb(55, 65, 81))
                .Table(rows, PdfAlign.Left, style)
                .Paragraph(p => p.Text("End of lists and tables sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80))), new PdfCreateOptions {
                    HeaderFormat = "OfficeIMO.Pdf lists and tables",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Lists and Tables", author: "OfficeIMO")
                .Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
