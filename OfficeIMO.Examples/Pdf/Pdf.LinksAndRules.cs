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
                [(1, 0)] = "https://officeimo.com/",
                [(1, 1)] = "https://officeimo.com/",
                [(1, 2)] = "https://officeimo.com/docs",
                [(2, 0)] = "https://github.com/EvotecIT/OfficeIMO",
                [(2, 1)] = "https://github.com/EvotecIT/OfficeIMO",
                [(2, 2)] = "https://github.com/EvotecIT/OfficeIMO/issues"
            };

            var tableStyle = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            tableStyle.CellPaddingX = 6;
            tableStyle.CellPaddingY = 5;
            tableStyle.CellPaddingTop = 5;
            tableStyle.CellPaddingBottom = 5;
            tableStyle.Caption = "Table 1. Linked resources";
            tableStyle.CaptionFontSize = 8.5;
            tableStyle.CaptionSpacingAfter = 5;
            tableStyle.SpacingBefore = 6;
            tableStyle.SpacingAfter = 12;

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Links & Rules Demo", PdfAlign.Left, linkUri: "https://github.com/EvotecIT/OfficeIMO")
                .Paragraph(p => p
                    .Text("Visit ")
                    .Link("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", PdfColor.FromRgb(20, 90, 180))
                    .Text(" and the ")
                    .Link("project website", "https://officeimo.com/", PdfColor.FromRgb(20, 90, 180))
                    .Text(" for more details."))
                .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
                .TableWithLinks(rows, links, PdfAlign.Left, tableStyle)
                .PanelParagraph(
                    p => p
                        .Text("You can also place links ")
                        .Link("inside panels", "https://officeimo.com/", PdfColor.FromRgb(20, 90, 180))
                        .Text("."),
                    new PdfPanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        BorderWidth = 0.5,
                        PaddingX = 9,
                        PaddingY = 7
                    })), new PdfCreateOptions {
                    HeaderFormat = "OfficeIMO.Pdf links and rules",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Links and Rules", author: "OfficeIMO")
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
