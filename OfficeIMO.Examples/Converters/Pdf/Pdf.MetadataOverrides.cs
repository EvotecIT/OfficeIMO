using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithMetadataOverrides(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with overridden metadata and exporting to PDF");
            string docPath = Path.Combine(folderPath, "PdfWithMetadataOverrides.docx");
            string pdfPath = Path.Combine(folderPath, "PdfWithMetadataOverrides.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.BuiltinDocumentProperties.Title = "Original Title";
                document.BuiltinDocumentProperties.Creator = "Original Author";
                document.BuiltinDocumentProperties.Subject = "Original Subject";
                document.BuiltinDocumentProperties.Keywords = "orig1, orig2";
                document.AddParagraph("Test");
                document.Save();
                var options = new PdfSaveOptions {
                    Title = "Pdf Title",
                    Author = "Pdf Author",
                    Subject = "Pdf Subject",
                    Keywords = "keyword1, keyword2"
                };
                document.SaveAsPdf(pdfPath, options);
            }
        }
    }
}
