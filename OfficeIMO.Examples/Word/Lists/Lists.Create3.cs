using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists6(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with lists - Document 6");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists6.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList = document.AddList(WordListStyle.Headings111);
                wordList.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList.AddItem("Text 2.1", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                wordList.AddItem("Text 2.2", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                // here we set another list element but we also change it using standard paragraph change
                paragraph = wordList.AddItem("Text 3");
                paragraph.Bold = true;
                paragraph.SetItalic();

                paragraph = document.AddParagraph("This is second list").SetColor(SixLabors.ImageSharp.Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList1 = document.AddList(WordListStyle.HeadingIA1);
                wordList1.AddItem("Temp 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList1.AddItem("Temp 2.1", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                wordList1.AddItem("Temp 2.2", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                wordList1.AddItem("Temp 2.3", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                wordList1.AddItem("Temp 2.3.4", 2).SetColor(SixLabors.ImageSharp.Color.Brown).Remove();
                wordList1.ListItems[1].Remove();
                wordList1.AddItem("Temp 3");

                document.Lists[0].AddItem("Added 2.3.5", 3).SetColor(SixLabors.ImageSharp.Color.DimGrey);


                Console.WriteLine("Lists count - before adding section: " + document.Lists.Count);

                var section = document.AddSection();
                section.PageSettings.Orientation = PageOrientationValues.Landscape;

                Console.WriteLine(document.Sections[0].PageSettings.Orientation);
                Console.WriteLine(document.Sections[1].PageSettings.Orientation);

                WordList wordList2 = document.AddList(WordListStyle.Headings111);
                wordList2.AddItem("Section 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList2.AddItem("Section 2.1", 1).SetColor(SixLabors.ImageSharp.Color.Brown);
                wordList2.AddItem("Section 2.2", 1).SetColor(SixLabors.ImageSharp.Color.Brown);

                Console.WriteLine("Lists count - after adding section and 1 list: " + document.Lists.Count);
                Console.WriteLine("Lists count - section 0: " + document.Sections[0].Lists.Count);
                Console.WriteLine("Lists count - section 1: " + document.Sections[1].Lists.Count);

                document.Save(openWord);
            }
        }
    }
}
