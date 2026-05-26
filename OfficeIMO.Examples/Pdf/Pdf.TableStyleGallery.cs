using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class TableStyleGalleryPdf {
        public static void Example_Pdf_TableStyleGallery(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.TableStyleGallery.pdf");
            var rows = new[] {
                new[] { "Signal", "State", "Notes" },
                new[] { "Header", "Repeated", "The first row should stay readable without relying on a domain-specific preset." },
                new[] { "Flow", "Generic", "Borders, row separators, and spacing should reveal the preset shape at raster level." }
            };

            PdfDoc doc = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9.5,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf Word-like table styles",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Table Style Gallery", author: "OfficeIMO")
                .H1("Table Style Gallery", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("Generic Word-style table names rendered by OfficeIMO.Pdf without invoice or report-specific behavior."));

            foreach (string styleName in TableStyles.SupportedWordStyleNames) {
                PdfTableStyle style = TableStyles.FromWordTableStyle(styleName);
                style.Caption = styleName;
                style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
                style.CaptionFontSize = 8.5;
                style.CaptionSpacingAfter = 4;
                style.SpacingBefore = 6;
                style.SpacingAfter = 6;
                style.ColumnWidthPoints = new List<double?> { 76, 70, 300 };
                style.AutoFitColumns = false;
                style.Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Center,
                    PdfColumnAlign.Left
                };

                doc.Table(rows, PdfAlign.Left, style);
            }

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
