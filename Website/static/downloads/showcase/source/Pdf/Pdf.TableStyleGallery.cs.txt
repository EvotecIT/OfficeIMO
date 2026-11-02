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
                new[] { "Flow", "Generic", "Borders, row separators and spacing distinguish each preset when viewed or printed." }
            };

            PdfDocument doc = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Table Style Gallery", PdfAlign.Left)
                .Paragraph(p => p.Text("Generic Word-style table names rendered by OfficeIMO.Pdf without invoice or report-specific behavior."))), new PdfCreateOptions {
                    HeaderFormat = "OfficeIMO.Pdf Word-like table styles",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Table Style Gallery", author: "OfficeIMO");

            foreach (string styleName in PdfTableStyles.SupportedWordStyleNames) {
                PdfTableStyle style = PdfTableStyles.FromWordTableStyle(styleName);
                style.Caption = styleName;
                style.CaptionColor = OfficeColor.FromRgb(80, 90, 100);
                style.CaptionFontSize = OfficeLength.Points(8.5);
                style.CaptionSpacingAfter = OfficeLength.Points(4);
                style.SpacingBefore = OfficeLength.Points(6);
                style.SpacingAfter = OfficeLength.Points(6);
                style.ColumnWidthWeights = new List<double> { 76, 70, 300 };
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
