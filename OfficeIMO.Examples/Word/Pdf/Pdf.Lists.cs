using System;
using System.IO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfWithLists(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document with nested lists as PDF");

            string docPath = Path.Combine(folderPath, "SaveAsPdfWithLists.docx");
            string pdfPath = Path.Combine(folderPath, "SaveAsPdfWithLists.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                WordList numberList = document.AddList(WordListStyle.Headings111);
                WordParagraph first = numberList.AddItem("Item 1");
                WordParagraph second = numberList.AddItem("Item 1.1");
                second.ListItemLevel = 1;
                WordParagraph third = numberList.AddItem("Item 1.1.1");
                third.ListItemLevel = 2;

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Bullet 1");
                WordParagraph nested = bulletList.AddItem("Bullet 1.1");
                nested.ListItemLevel = 1;

                document.Save();
                document.SaveAsPdf(pdfPath);
            }
        }
    }
}

