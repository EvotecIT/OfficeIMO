using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyleCheatsheet {
        public static void Example_Pdf_StyleCheatsheet(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyleCheatsheet.pdf");

            var doc = PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Style Cheatsheet", OfficeAlignment.Left)
                .AddPanelParagraph(
                    p => p.AddText("A compact visual sample for rich text, color, underline, and alignment behavior."),
                    new PdfPanelStyle {
                        Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                        BorderColor = global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207),
                        PaddingX = global::OfficeIMO.OfficeLength.Points(9),
                        PaddingY = global::OfficeIMO.OfficeLength.Points(7)
                    })
                .AddParagraph(p => {
                    p.AddText("Normal ");
                    p.AddText("Bold ").SetBold();
                    p.AddText("Italic ").SetItalic();
                    p.AddText("Bold").SetBold();
                    p.AddText(" Italic ").SetItalic();
                    p.AddText("Underline ").SetUnderline();
                })
                .AddParagraph(p => {
                    p.AddText("Colors: ");
                    p.AddText("Red ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(200,0,0));
                    p.AddText("Blue ").SetColor(global::OfficeIMO.OfficeColor.FromRgb(20,90,180));
                    p.AddText("Green").SetColor(global::OfficeIMO.OfficeColor.FromRgb(0,128,0));
                })
                .AddParagraph(p => {
                    p.AddText("Combinations: ");
                    p.AddText("Bold ").SetBold();
                    p.AddText("Italic ").SetItalic();
                    p.AddText("Underlined ").SetUnderline();
                    p.AddText("Bold ").SetBold();
                    p.AddText("Italic ").SetItalic();
                    p.AddText("Underlined ").SetUnderline();
                })
                .AddParagraph(p => {
                    p.AddText("Stateful toggles: ");
                    p.AddText("bold on ").SetBold();
                    p.AddText("bold off ");
                    p.AddText("italic on ").SetItalic();
                    p.AddText("italic off ");
                    p.AddText("ul on ").SetUnderline();
                    p.AddText("ul off");
                })
                .HR(global::OfficeIMO.OfficeLength.Points(0.8), global::OfficeIMO.OfficeColor.FromRgb(183, 194, 207), global::OfficeIMO.OfficeLength.Points(8), global::OfficeIMO.OfficeLength.Points(8))
                .AddParagraph(p => p.AddText("Center aligned line"), OfficeAlignment.Center)
                .AddParagraph(p => p.AddText("Right aligned line"), OfficeAlignment.Right)), new PdfCreateOptions {
                    PageSetup = {
                        Header = { Format = "OfficeIMO.Pdf style cheatsheet", Alignment = OfficeAlignment.Left, Show = true },
                        Footer = { Alignment = OfficeAlignment.Right, Show = true }
                    }
                }).Meta(title: "OfficeIMO.Pdf Style Cheatsheet", author: "OfficeIMO");

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
