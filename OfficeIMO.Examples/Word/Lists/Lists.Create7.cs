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
                wordList1.AddItem("Text 2.1", 1);

                WordList wordList2 = document.AddList(WordListStyle.Headings111, true);
                wordList2.RestartNumbering = true;
                wordList2.AddItem("Section 1");
                wordList2.AddItem("Section 2.1", 1);

                WordList wordList3 = document.AddList(WordListStyle.Headings111, true);
                wordList3.RestartNumbering = true;
                wordList3.AddItem("Section 1");
                wordList3.AddItem("Section 2.1", 1);

                WordList wordList4 = document.AddList(WordListStyle.Headings111, true);
                wordList4.RestartNumbering = true;
                wordList4.AddItem("Section 1");
                wordList4.AddItem("Section 2.1", 1);

                WordList wordList5 = document.AddList(WordListStyle.Headings111, true);
                wordList5.RestartNumbering = true;
                wordList5.AddItem("Section 1");
                wordList5.AddItem("Section 2.1", 1);

                document.Save(openWord);
            }
        }
    }
}
