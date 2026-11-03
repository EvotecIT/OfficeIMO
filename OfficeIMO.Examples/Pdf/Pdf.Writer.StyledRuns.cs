using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyledRuns {
        public static void Example_Pdf_StyledRuns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyledRuns.pdf");
            var options = new PdfCreateOptions {
                HeaderFormat = "OfficeIMO.Pdf styled runs",
                HeaderAlign = OfficeAlignment.Left,
                ShowHeader = true,
                FooterAlign = OfficeAlignment.Right,
                ShowPageNumbers = true
            };

            var doc = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Styled Runs", OfficeAlignment.Left)
                .AddParagraph(p => p.AddText("A compact visual sample for inline font style, color, underline, strike-through, and stateful run toggles."))
                .AddPanelParagraph(
                    p => {
                        p.AddText("Inline styles").SetBold();
                        p.AddText(" remain readable in ordinary business reports.");
                    },
                    new PdfPanelStyle {
                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                        PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                        PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                    })
                .AddParagraph(p => {
                    p.AddText("You can mix ");
                    p.AddText("bold ").SetBold();
                    p.AddText("italic ").SetItalic();
                    p.AddText("bold italic ").SetBold().SetItalic();
                    p.AddText("underlined ").SetUnderline();
                    p.AddText("obsolete ").SetStrikethrough();
                    p.AddText("and ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80));
                    p.AddText("colors.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(8, 28, 120));
                },
                    style: new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(14) })
                .AddParagraph(p => {
                    p.AddText("Underline respects color: ");
                    p.AddText("red").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(200, 0, 0));
                    p.AddText(", ");
                    p.AddText("blue").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                    p.AddText(", and ");
                    p.AddText("green").SetUnderline().SetColor(global::OfficeIMO.OfficeColor.FromRgb(0, 128, 0));
                    p.AddText(".");
                })
                .AddParagraph(p => {
                    p.AddText("Decoration variants: ");
                    p.AddText("double").SetUnderline(OfficeIMO.Drawing.OfficeTextDecorationStyle.Double);
                    p.AddText(", ");
                    p.AddText("dotted").SetUnderline(OfficeIMO.Drawing.OfficeTextDecorationStyle.Dotted);
                    p.AddText(", ");
                    p.AddText("wavy strike").SetStrikethrough(OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy);
                    p.AddText("; H");
                    p.AddText("2").SetSubscript();
                    p.AddText("O and x");
                    p.AddText("2").SetSuperscript();
                    p.AddText(".");
                })
                .AddParagraph(p => p.AddRuns(new[] {
                    new PdfTextRun("stored casing", bold: true, color: global::OfficeIMO.OfficeColor.FromRgb(112, 48, 160))
                        .WithTextCase(OfficeIMO.Drawing.OfficeTextCase.TitleCase)
                }))
                .AddParagraph(p => {
                    p.AddText("Stateful color toggles: ");
                    p.AddText("critical ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(185, 28, 28));
                    p.AddText("informational ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180));
                    p.AddText("healthy ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52));
                    p.AddText("normal.").SetColor(global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55));
                })
                .AddParagraph(p => p.AddText("End of styled runs sample."), OfficeAlignment.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))), options)
                .Meta(title: "OfficeIMO.Pdf Styled Runs", author: "OfficeIMO");
            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
