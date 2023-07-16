using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists7(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with lists - Document 7");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists10.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordList wordList1 = document.AddList(WordListStyle.Headings111, true);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 1.1", 1);

                WordList wordList2 = document.AddList(WordListStyle.Headings111, true);
                Console.WriteLine("List 2 - Restart numbering: " + wordList2.RestartNumbering);
                wordList2.AddItem("Section 2");
                wordList2.AddItem("Section 2.1", 1);

                WordList wordList3 = document.AddList(WordListStyle.Headings111, true);
                Console.WriteLine("List 3 - Restart numbering: " + wordList3.RestartNumbering);
                wordList3.RestartNumbering = true;
                Console.WriteLine("List 3 - Restart numbering after change: " + wordList3.RestartNumbering);
                wordList3.AddItem("Section 1");
                wordList3.AddItem("Section 1.1", 1);

                WordList wordList4 = document.AddList(WordListStyle.Headings111, true);
                //wordList4.RestartNumbering = true;
                wordList4.AddItem("Section 2");
                wordList4.AddItem("Section 2.1", 1);

                WordList wordList5 = document.AddList(WordListStyle.Headings111, true);
                //wordList5.RestartNumbering = true;
                wordList5.AddItem("Section 3");
                wordList5.AddItem("Section 3.1", 1);

                WordList wordList6 = document.AddList(WordListStyle.Headings111);
                wordList1.AddItem("Text 4");
                wordList1.AddItem("Text 4.1", 1);

                document.Save(openWord);
            }
        }
    }
}
