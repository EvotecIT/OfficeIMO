using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists8(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with lists - Document 8");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists11.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                // add list and nest a list
                WordList wordList1 = document.AddList(WordListStyle.Headings111, false);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 1.1");
                wordList1.AddItem("Text 1.2");
                wordList1.AddItem("Text 1.3");

                document.AddParagraph("Second List");
                document.AddParagraph();

                WordList wordList2 = document.AddList(WordListStyle.Headings111, false);
                wordList2.AddItem("Text 2");
                wordList2.AddItem("Text 2.1");
                wordList2.AddItem("Text 2.2");
                wordList2.AddItem("Text 2.3");
                wordList2.AddItem("Text 2.4");

                document.AddParagraph("Third List");
                document.AddParagraph();

                WordList wordList3 = document.AddList(WordListStyle.Headings111, false);
                wordList3.AddItem("Text 3");
                wordList3.AddItem("Text 3.1");
                wordList3.AddItem("Text 3.2");
                wordList3.AddItem("Text 3.3");
                wordList3.AddItem("Text 3.4");


                document.Save(openWord);
            }
        }
    }
}
