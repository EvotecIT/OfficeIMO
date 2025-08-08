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
                document.BuiltinDocumentProperties.Title = "Pdf Title";
                document.BuiltinDocumentProperties.Creator = "Pdf Author";
                document.BuiltinDocumentProperties.Subject = "Pdf Subject";
                document.BuiltinDocumentProperties.Keywords = "keyword1, keyword2";
                document.AddParagraph("Test");
                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
