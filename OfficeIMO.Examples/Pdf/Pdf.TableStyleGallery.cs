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

            PdfDocument doc = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Table Style Gallery", PdfAlign.Left, global::OfficeIMO.OfficeColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("Generic Word-style table names rendered by OfficeIMO.Pdf without invoice or report-specific behavior."))), new PdfCreateOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9.5),
                    DefaultTextColor = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = global::OfficeIMO.OfficeLength.Points(8),
                    HeaderFormat = "OfficeIMO.Pdf Word-like table styles",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = global::OfficeIMO.OfficeLength.Points(8),
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Table Style Gallery", author: "OfficeIMO");

            foreach (string styleName in PdfTableStyles.SupportedWordStyleNames) {
                PdfTableStyle style = PdfTableStyles.FromWordTableStyle(styleName);
                style.Caption = styleName;
                style.CaptionColor = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100);
                style.CaptionFontSize = global::OfficeIMO.OfficeLength.Points(8.5);
                style.CaptionSpacingAfter = global::OfficeIMO.OfficeLength.Points(4);
                style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
                style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(6);
                style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(76), global::OfficeIMO.OfficeLength.Points(70), global::OfficeIMO.OfficeLength.Points(300) };
                style.AutoFitColumns = false;
                style.Alignments = new List<PdfColumnAlign> {
                    PdfColumnAlign.Left,
                    PdfColumnAlign.Center,
                    PdfColumnAlign.Left
                };

                doc.Compose(pdf => pdf.Content(content => content.Table(rows, PdfAlign.Left, style)));
            }

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
