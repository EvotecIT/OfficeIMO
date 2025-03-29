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
        internal static void Example_BasicLists(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with lists");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList = document.AddList(WordListStyle.Headings111);
                wordList.AddItem("Text 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList.AddItem("Text 1.1", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 1.2", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 1.3", 1).SetColor(Color.Brown);
                wordList.AddItem("Text 1.3.1", 2).SetColor(Color.Brown);
                // here we set another list element but we also change it using standard paragraph change
                paragraph = wordList.AddItem("Text 2");
                paragraph.Bold = true;
                paragraph.SetItalic();

                paragraph = document.AddParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList1 = document.AddList(WordListStyle.HeadingIA1);
                wordList1.AddItem("Temp 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList1.AddItem("Temp 2.1", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.2", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.3", 1).SetColor(Color.Brown);
                wordList1.AddItem("Temp 2.3.4", 2).SetColor(Color.Brown).Remove();
                wordList1.ListItems[1].Remove();
                paragraph = wordList1.AddItem("Temp 3");

                paragraph = document.AddParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(WordListStyle.BulletedChars);
                wordList2.AddItem("Oops 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList2.AddItem("Oops 2.1", 1).SetColor(Color.Brown);
                wordList2.AddItem("Oops 2.2", 1).SetColor(Color.Brown);
                wordList2.AddItem("Oops 2.3", 1).SetColor(Color.Brown);
                wordList2.AddItem("Oops 2.3.4", 2).SetColor(Color.Brown);

                Console.WriteLine("Number of lists - section 0: " + document.Sections[0].Lists.Count);
                Console.WriteLine("Number of lists - all: " + document.Lists.Count);

                document.AddSection();
                document.Sections[1].PageSettings.Orientation = PageOrientationValues.Landscape;


                Console.WriteLine("Number of lists - section 1: " + document.Sections[1].Lists.Count);
                Console.WriteLine("Number of lists - all: " + document.Lists.Count);


                paragraph = document.AddParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                WordList wordList3 = document.AddList(WordListStyle.Heading1ai);
                wordList3.AddItem("4th 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList3.AddItem("4th 2.1", 1).SetColor(Color.Brown);
                wordList3.AddItem("4th 2.2", 1).SetColor(Color.Brown);
                wordList3.AddItem("4th 2.3", 1).SetColor(Color.Brown);
                wordList3.AddItem("4th 2.3.4", 2).SetColor(Color.Brown);

                paragraph = document.AddParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(WordListStyle.Headings111Shifted);
                wordList4.AddItem("5th 1").SetCapsStyle(CapsStyle.SmallCaps);
                wordList4.AddItem("5th 2.1", 1).SetColor(Color.Brown);
                wordList4.AddItem("5th 2.2", 1).SetColor(Color.Brown);
                wordList4.AddItem("5th 2.3", 1).SetColor(Color.Brown);
                wordList4.AddItem("5th 2.3.4", 2).SetColor(Color.Brown);

                Console.WriteLine("Number of lists - section 1 (after adding 2): " + document.Sections[1].Lists.Count);
                Console.WriteLine("Number of lists - all (after adding 2): " + document.Lists.Count);

                document.Lists[3].ListItems[2].Text = "Overwrite Text 2.2";
                document.Lists[4].ListItems[2].Text = "Overwrite Text 2.2";

                paragraph = document.AddParagraph("This is continuation, we now add 2 items to two lists ").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                document.Lists[3].AddItem("Added 2.3.5 to list number 4", 2).SetColor(Color.DimGrey);
                document.Lists[2].AddItem("Added 2.3.5 to list number 3", 2).SetColor(Color.DimGrey);

                document.AddBreak();

                wordList.AddItem("Text 2.1.1", 2).SetColor(Color.Brown);
                wordList.AddItem("Text 2.1.2", 2).SetColor(Color.Brown);

                Console.WriteLine("Number of lists - section 1 (after updating existing lists): " + document.Sections[1].Lists.Count);
                Console.WriteLine("Number of lists - all (after updating existing lists): " + document.Lists.Count);

                document.Save(openWord);
            }
        }
    }
}
