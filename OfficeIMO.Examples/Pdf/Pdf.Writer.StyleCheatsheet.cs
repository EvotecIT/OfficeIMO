using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyleCheatsheet {
        public static void Example_Pdf_StyleCheatsheet(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyleCheatsheet.pdf");

            var doc = PdfDoc.Create(new PdfOptions { DefaultTextColor = PdfColor.FromRgb(40, 40, 40) })
                .H1("Style Cheatsheet", PdfAlign.Center)

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

                .Paragraph(p => p.Text(" "), PdfAlign.Center)
                .Paragraph(p => p.Text("Center aligned line"), PdfAlign.Center)
                .Paragraph(p => p.Text("Right aligned line"), PdfAlign.Right)
                ;

            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
