using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageEmbedderHelper(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with ImageEmbedder helper");
            string filePath = Path.Combine(folderPath, "ImageEmbedder.docx");

            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document, true)) {
                MainDocumentPart mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                string imagePath = Path.Combine("Assets", "OfficeIMO.png");
                Paragraph p = new Paragraph();
                p.Append(ImageEmbedder.CreateImageRun(mainPart, imagePath));
                mainPart.Document.Body.Append(p);
                mainPart.Document.Save();
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
