using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;
namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithMetadata(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with metadata and exporting to PDF");
            string docPath = Path.Combine(folderPath, "PdfWithMetadata.docx");
            string pdfPath = Path.Combine(folderPath, "PdfWithMetadata.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Properties.Title = "Pdf Title";
                document.Properties.Creator = "Pdf Author";
                document.Properties.Subject = "Pdf Subject";
                document.Properties.Keywords = "keyword1, keyword2";
                document.AddParagraph("Test");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
