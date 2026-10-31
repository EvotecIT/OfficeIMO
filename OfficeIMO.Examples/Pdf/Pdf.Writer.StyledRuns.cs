using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyledRuns {
        public static void Example_Pdf_StyledRuns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyledRuns.pdf");
            var options = new PdfCreateOptions {
                HeaderFormat = "OfficeIMO.Pdf styled runs",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            };

            var doc = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Styled Runs", PdfAlign.Left)
                .Paragraph(p => p.Text("A compact visual sample for inline font style, color, underline, strike-through, and stateful run toggles."))
                .PanelParagraph(
                    p => p
                        .Bold("Inline styles")
                        .Text(" remain readable in ordinary business reports."),
                    new PdfPanelStyle {
                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                        PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                        PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                    })
                .Paragraph(p => p
                    .Text("You can mix ")
                    .Bold("bold ")
                    .Italic("italic ")
                    .Bold(true).Italic(true).Text("bold italic ").Italic(false).Bold(false)
                    .Underlined("underlined ")
                    .Strikethrough("obsolete ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80)).Text("and ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(8, 28, 120)).Text("colors."),
                    style: new PdfParagraphStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Points(14) })
                .Paragraph(p => p
                    .Text("Underline respects color: ")
                    .Underlined("red", global::OfficeIMO.OfficeColor.FromRgb(200, 0, 0))
                    .Text(", ")
                    .Underlined("blue", global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180))
                    .Text(", and ")
                    .Underlined("green", global::OfficeIMO.OfficeColor.FromRgb(0, 128, 0))
                    .Text("."))
                .Paragraph(p => p
                    .Text("Decoration variants: ")
                    .Underlined("double", OfficeIMO.Drawing.OfficeTextDecorationStyle.Double)
                    .Text(", ")
                    .Underlined("dotted", OfficeIMO.Drawing.OfficeTextDecorationStyle.Dotted)
                    .Text(", ")
                    .Strikethrough("wavy strike", OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy)
                    .Text("; H")
                    .Subscript("2")
                    .Text("O and x")
                    .Superscript("2")
                    .Text("."))
                .Paragraph(p => p.Runs(new[] {
                    new PdfTextRun("stored casing", bold: true, color: global::OfficeIMO.OfficeColor.FromRgb(112, 48, 160))
                        .WithTextCase(OfficeIMO.Drawing.OfficeTextCase.TitleCase)
                }))
                .Paragraph(p => p
                    .Text("Stateful color toggles: ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(185, 28, 28)).Text("critical ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(20, 90, 180)).Text("informational ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(22, 101, 52)).Text("healthy ")
                    .Color(global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55)).Text("normal."))
                .Paragraph(p => p.Text("End of styled runs sample."), PdfAlign.Right, global::OfficeIMO.OfficeColor.FromRgb(80, 80, 80))), options)
                .Meta(title: "OfficeIMO.Pdf Styled Runs", author: "OfficeIMO");
            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
