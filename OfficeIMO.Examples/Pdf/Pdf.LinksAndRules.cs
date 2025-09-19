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
                HeaderFill = PdfColor.LightGray,
                RowStripeFill = PdfColor.FromRgb(248, 248, 248),
                BorderColor = PdfColor.FromRgb(210, 210, 210),
                BorderWidth = 0.5,
                CellPaddingX = 4,
                CellPaddingY = 2
            };

            PdfDoc.Create()
                .H1("Links & Rules Demo", PdfAlign.Center, PdfColor.FromRgb(8, 28, 120), linkUri: "https://github.com/EvotecIT/OfficeIMO")
                .Paragraph(p => p
                    .Text("Visit ")
                    .Link("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", PdfColor.FromRgb(20, 90, 180))
                    .Text(" and the ")
                    .Link("project website", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                    .Text(" for more details."))
                .HR(0.8, PdfColor.Gray, 8, 8)
                .TableWithLinks(rows, links, PdfAlign.Left, tableStyle)
                .HR(0.8, PdfColor.Gray, 8, 8)
                .PanelParagraph(
                    p => p
                        .Text("You can also place links ")
                        .Link("inside panels", "https://officeimo.net/", PdfColor.FromRgb(20, 90, 180))
                        .Text("."),
                    new PanelStyle { Background = new PdfColor(0.95, 0.95, 0.98), BorderColor = PdfColor.FromRgb(210, 210, 210), BorderWidth = 0.5, PaddingY = 6 }
                )
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
