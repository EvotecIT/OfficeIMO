using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_PdfPageNumbers(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with page number options");
            string docPath = Path.Combine(folderPath, "PdfPageNumbers.docx");
            string pdfNoNumbers = Path.Combine(folderPath, "PdfWithoutNumbers.pdf");
            string pdfCustomNumbers = Path.Combine(folderPath, "PdfCustomNumbers.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();
                document.SaveAsPdf(pdfNoNumbers, new PdfSaveOptions { IncludePageNumbers = false });
                document.SaveAsPdf(pdfCustomNumbers, new PdfSaveOptions { PageNumberFormat = "Page {current} of {total}" });
            }
        }
    }
}
