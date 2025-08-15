using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_PdfFontFamily(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom font");
            string docPath = Path.Combine(folderPath, "PdfFontFamily.docx");
            string pdfPath = Path.Combine(folderPath, "PdfFontFamily.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    FontFamily = "Times New Roman"
                });
            }
        }
    }
}
