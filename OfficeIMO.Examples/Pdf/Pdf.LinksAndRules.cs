using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class LinksAndRules {
        public static void Example_Pdf_LinksAndRules(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.LinksAndRules.pdf");

            var rows = new[] {
                new [] { "Site", "Label", "Notes" },
                new [] { "OfficeIMO", "Homepage", "Docs" },
                new [] { "GitHub", "Repo", "Issues" }
            };

            var links = new Dictionary<(int Row, int Col), string> {
                [(1, 0)] = "https://officeimo.net/",
                [(1, 1)] = "https://officeimo.net/",
                [(1, 2)] = "https://officeimo.net/docs",
                [(2, 0)] = "https://github.com/EvotecIT/OfficeIMO",
                [(2, 1)] = "https://github.com/EvotecIT/OfficeIMO",
                [(2, 2)] = "https://github.com/EvotecIT/OfficeIMO/issues"
            };

            var tableStyle = new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(32, 76, 120),
                HeaderTextColor = PdfColor.White,
                TextColor = PdfColor.FromRgb(31, 41, 55),
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                BorderColor = PdfColor.FromRgb(210, 218, 226),
                BorderWidth = 0.5,
                CellPaddingX = 6,
                CellPaddingY = 5,
                Caption = "Table 1. Linked resources",
                CaptionColor = PdfColor.FromRgb(80, 90, 100),
                CaptionFontSize = 8.5,
                CaptionSpacingAfter = 5,
                SpacingBefore = 6,
                SpacingAfter = 12
            };

            PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf links and rules",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Links and Rules", author: "OfficeIMO")
                .H1("Links & Rules Demo", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85), linkUri: "https://github.com/EvotecIT/OfficeIMO")
                .Paragraph(p => p
                    .Text("Visit ")
                    .Link("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", PdfColor.FromRgb(20, 90, 180))
                    .Text(" and the ")
                    .Link("project website", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                    .Text(" for more details."))
                .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
                .TableWithLinks(rows, links, PdfAlign.Left, tableStyle)
                .PanelParagraph(
                    p => p
                        .Text("You can also place links ")
                        .Link("inside panels", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                        .Text("."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        BorderWidth = 0.5,
                        PaddingX = 9,
                        PaddingY = 7
                    }
                )
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
