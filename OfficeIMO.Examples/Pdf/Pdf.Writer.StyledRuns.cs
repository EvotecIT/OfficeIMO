using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterStyledRuns {
        public static void Example_Pdf_StyledRuns(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.StyledRuns.pdf");
            var doc = PdfDoc.Create()
                .H1("Inline Styles Demo", PdfAlign.Center)
                .Paragraph(p => p
                    .Text("You can mix ")
                    .Bold("bold ")
                    .Italic("italic ")
                    .Bold("bold italic ") // with current engine, bold is applied; add .Italic if desired
                    .Underlined("underlined ")
                    .Color(PdfColor.FromRgb(80, 80, 80)).Text("and ")
                    .Color(PdfColor.FromRgb(8, 28, 120)).Text("colors."))
                .Paragraph(p => p.Text(" "))
                .Paragraph(p => p
                    .Text("Underline respects color: ")
                    .Underlined("red", PdfColor.FromRgb(200, 0, 0))
                    .Text(", ")
                    .Underlined("blue", PdfColor.FromRgb(20, 90, 180))
                    .Text("."));
            doc.Save(path);
            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}
