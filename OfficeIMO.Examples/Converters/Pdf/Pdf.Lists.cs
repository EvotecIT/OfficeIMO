using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveLists(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with lists and exporting to PDF");
            string docPath = Path.Combine(folderPath, "ExportListsToPdf.docx");
            string pdfPath = Path.Combine(folderPath, "ExportListsToPdf.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numbered = document.AddList(WordListStyle.Headings111);
                numbered.AddItem("First");
                numbered.AddItem("Second");
                numbered.AddItem("Second - Nested", 1);
                numbered.AddItem("Third");

                WordList bullets = document.AddList(WordListStyle.Bulleted);
                bullets.AddItem("Alpha");
                bullets.AddItem("Beta");
                bullets.AddItem("Beta - Nested", 1);
                bullets.AddItem("Gamma");

                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}
