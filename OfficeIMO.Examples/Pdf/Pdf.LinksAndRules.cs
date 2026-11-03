using OfficeIMO.Pdf;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class LinksAndRules {
        public static void Example_Pdf_LinksAndRules(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.LinksAndRules.pdf");

            var rows = new[] {
                new[] { PdfTableCell.TextCell("Site"), PdfTableCell.TextCell("Label"), PdfTableCell.TextCell("Notes") },
                new[] { PdfTableCell.TextCell("OfficeIMO", "https://officeimo.com/"), PdfTableCell.TextCell("Homepage", "https://officeimo.com/"), PdfTableCell.TextCell("Docs", "https://officeimo.com/docs") },
                new[] { PdfTableCell.TextCell("GitHub", "https://github.com/EvotecIT/OfficeIMO"), PdfTableCell.TextCell("Repo", "https://github.com/EvotecIT/OfficeIMO"), PdfTableCell.TextCell("Issues", "https://github.com/EvotecIT/OfficeIMO/issues") }
            };

            var tableStyle = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            tableStyle.CellPaddingX = OfficeLength.Points(6);
            tableStyle.CellPaddingY = OfficeLength.Points(5);
            tableStyle.CellPaddingTop = OfficeLength.Points(5);
            tableStyle.CellPaddingBottom = OfficeLength.Points(5);
            tableStyle.Caption = "Table 1. Linked resources";
            tableStyle.CaptionFontSize = OfficeLength.Points(8.5);
            tableStyle.CaptionSpacingAfter = OfficeLength.Points(5);
            tableStyle.SpacingBefore = OfficeLength.Points(6);
            tableStyle.SpacingAfter = OfficeLength.Points(12);

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Links & Rules Demo", OfficeAlignment.Left, linkUri: "https://github.com/EvotecIT/OfficeIMO")
                .AddParagraph(p => {
                    p.AddText("Visit ");
                    p.AddLink("OfficeIMO GitHub", "https://github.com/EvotecIT/OfficeIMO", global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                    p.AddText(" and the ");
                    p.AddLink("project website", "https://officeimo.com/", OfficeColor.FromRgb(20, 90, 180));
                    p.AddText(" for more details.");
                })
                .HR(global::OfficeIMO.OfficeLength.Points(0.8), global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), global::OfficeIMO.OfficeLength.Points(8), global::OfficeIMO.OfficeLength.Points(8))
                .AddTable(rows, OfficeAlignment.Left, tableStyle)
                .AddPanelParagraph(
                    p => {
                        p.AddText("You can also place links ");
                        p.AddLink("inside panels", "https://officeimo.com/", OfficeColor.FromRgb(20, 90, 180));
                        p.AddText(".");
                    },
                    new PdfPanelStyle {
                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                        BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
                        PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                        PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                    })), new PdfCreateOptions {
                    HeaderFormat = "OfficeIMO.Pdf links and rules",
                    HeaderAlign = OfficeAlignment.Left,
                    ShowHeader = true,
                    FooterAlign = OfficeAlignment.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Links and Rules", author: "OfficeIMO")
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
