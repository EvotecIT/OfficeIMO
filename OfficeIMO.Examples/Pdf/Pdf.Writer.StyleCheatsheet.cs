using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyleCheatsheet {
        public static void Example_Pdf_StyleCheatsheet(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyleCheatsheet.pdf");

            var doc = PdfDoc.Create(new PdfOptions {
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                    HeaderFont = PdfStandardFont.Helvetica,
                    HeaderFontSize = 8,
                    HeaderFormat = "OfficeIMO.Pdf style cheatsheet",
                    HeaderAlign = PdfAlign.Left,
                    ShowHeader = true,
                    FooterFont = PdfStandardFont.Helvetica,
                    FooterFontSize = 8,
                    FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                    FooterAlign = PdfAlign.Right,
                    ShowPageNumbers = true
                })
                .Meta(title: "OfficeIMO.Pdf Style Cheatsheet", author: "OfficeIMO")
                .H1("Style Cheatsheet", PdfAlign.Left, PdfColor.FromRgb(25, 55, 85))
                .PanelParagraph(
                    p => p.Text("A compact visual sample for rich text, color, underline, and alignment behavior."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(183, 194, 207),
                        PaddingX = 9,
                        PaddingY = 7
                    })

                .Paragraph(p => p
                    .Text("Normal ")
                    .Bold("Bold ")
                    .Italic("Italic ")
                    .Bold("Bold").Italic(" Italic ")
                    .Underlined("Underline ")
                )

                .Paragraph(p => p
                    .Text("Colors: ")
                    .Color(PdfColor.FromRgb(200,0,0)).Text("Red ")
                    .Color(PdfColor.FromRgb(20,90,180)).Text("Blue ")
                    .Color(PdfColor.FromRgb(0,128,0)).Text("Green")
                )

                .Paragraph(p => p
                    .Text("Combinations: ")
                    .Bold("Bold ")
                    .Italic("Italic ")
                    .Underlined("Underlined ")
                    .Bold("Bold ").Italic("Italic ").Underlined("Underlined ")
                )

                .Paragraph(p => p
                    .Text("Stateful toggles: ")
                    .Bold(true).Text("bold on ")
                    .Bold(false).Text("bold off ")
                    .Italic(true).Text("italic on ")
                    .Italic(false).Text("italic off ")
                    .Underline(true).Text("ul on ")
                    .Underline(false).Text("ul off")
                )

                .HR(0.8, PdfColor.FromRgb(183, 194, 207), 8, 8)
                .Paragraph(p => p.Text("Center aligned line"), PdfAlign.Center)
                .Paragraph(p => p.Text("Right aligned line"), PdfAlign.Right)
                ;

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
