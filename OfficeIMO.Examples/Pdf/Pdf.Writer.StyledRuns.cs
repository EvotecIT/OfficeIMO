using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyledRuns {
        public static void Example_Pdf_StyledRuns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyledRuns.pdf");
            var options = new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = "OfficeIMO.Pdf styled runs",
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true
            };

            var doc = PdfDoc.Create(options)
                .Meta(title: "OfficeIMO.Pdf Styled Runs", author: "OfficeIMO")
                .H1("Styled Runs", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .Paragraph(p => p.Text("A compact visual sample for inline font style, color, underline, strike-through, and stateful run toggles."))
                .PanelParagraph(
                    p => p
                        .Bold("Inline styles")
                        .Text(" should remain readable in normal business-report text, not only in synthetic text extraction checks."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        PaddingX = 9,
                        PaddingY = 7
                    })
                .Paragraph(p => p
                    .Text("You can mix ")
                    .Bold("bold ")
                    .Italic("italic ")
                    .Bold(true).Italic(true).Text("bold italic ").Italic(false).Bold(false)
                    .Underlined("underlined ")
                    .Strikethrough("obsolete ")
                    .Color(PdfColor.FromRgb(80, 80, 80)).Text("and ")
                    .Color(PdfColor.FromRgb(8, 28, 120)).Text("colors."),
                    style: new PdfParagraphStyle { SpacingBefore = 14 })
                .Paragraph(p => p
                    .Text("Underline respects color: ")
                    .Underlined("red", PdfColor.FromRgb(200, 0, 0))
                    .Text(", ")
                    .Underlined("blue", PdfColor.FromRgb(20, 90, 180))
                    .Text(", and ")
                    .Underlined("green", PdfColor.FromRgb(0, 128, 0))
                    .Text("."))
                .Paragraph(p => p
                    .Text("Stateful color toggles: ")
                    .Color(PdfColor.FromRgb(185, 28, 28)).Text("critical ")
                    .Color(PdfColor.FromRgb(20, 90, 180)).Text("informational ")
                    .Color(PdfColor.FromRgb(22, 101, 52)).Text("healthy ")
                    .Color(PdfColor.FromRgb(31, 41, 55)).Text("normal."))
                .Paragraph(p => p.Text("End of styled runs sample."), PdfAlign.Right, PdfColor.FromRgb(80, 80, 80));
            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
