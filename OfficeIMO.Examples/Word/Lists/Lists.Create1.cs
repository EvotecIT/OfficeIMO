using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists3(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists3.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("This is 1st list");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(WordListStyle.Headings111);
                wordList1.AddItem("Text 1 - List1");
                wordList1.AddItem("Text 2 - List1", 1);
                wordList1.AddItem("Text 3 - List1", 2);

                paragraph = document.AddParagraph("This is 2nd list");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList2 = document.AddList(WordListStyle.Headings111);
                wordList2.AddItem("Text 1");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 3rd list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList3 = document.AddList(WordListStyle.Bulleted);
                wordList3.AddItem("Text 7.1", 1);
                wordList3.AddItem("Text 7.2", 2);
                wordList3.AddItem("Text 7.3", 2);
                wordList3.AddItem("Text 7.4", 0);
                wordList3.AddItem("Text 7.5", 0);
                wordList3.AddItem("Text 7.6", 1);
                wordList3.AddItem("Text 7");

                paragraph = document.AddParagraph("This is 4th list").SetColor(Color.Aqua).SetUnderline(UnderlineValues.Double);
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList4 = document.AddList(WordListStyle.Bulleted);
                wordList4.AddItem("Text 8");
                wordList4.AddItem("Text 8.1", 1);
                wordList4.AddItem("Text 8.2", 2);
                wordList4.AddItem("Text 8.3", 2);
                wordList4.AddItem("Text 8.4", 0);
                wordList4.AddItem("Text 8.5", 0);
                wordList4.AddItem("Text 8.6", 1);

                Console.WriteLine("List count: " + document.Lists.Count); // "List count: 4

                document.Lists[0].Remove();

                Console.WriteLine("List count: " + document.Lists.Count); // "List count: 3

                document.Lists[0].Merge(document.Lists[1]);

                Console.WriteLine("List count: " + document.Lists.Count); // "List count: 2

                document.Save(openWord);
            }
        }
    }
}
