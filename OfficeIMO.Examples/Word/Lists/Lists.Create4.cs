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
        internal static void Example_BasicLists2(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(WordListStyle.ArticleSections);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is second list").SetColor(Color.OrangeRed).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(WordListStyle.Headings111);
                wordList2.AddItem("Temp 2");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);
                wordList2.AddItem("Text 3", 2);

                wordList2.ListItems[3].ListItemLevel = 0;

                paragraph = document.AddParagraph("This is third list").SetColor(Color.Blue).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(WordListStyle.HeadingIA1);
                wordList3.AddItem("Text 3");
                wordList3.AddItem("Text 2", 1);
                wordList3.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is fourth list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList4 = document.AddList(WordListStyle.Chapters); // Chapters support only level 0
                wordList4.AddItem("Text 1");
                wordList4.AddItem("Text 2");
                wordList4.AddItem("Text 3");

                paragraph = document.AddParagraph("This is five list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList5 = document.AddList(WordListStyle.BulletedChars);
                wordList5.AddItem("Text 5");
                wordList5.AddItem("Text 2", 1);
                wordList5.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 6th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList6 = document.AddList(WordListStyle.Heading1ai);
                wordList6.AddItem("Text 6");
                wordList6.AddItem("Text 2", 1);
                wordList6.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList7 = document.AddList(WordListStyle.Headings111Shifted);
                wordList7.AddItem("Text 7");
                wordList7.AddItem("Text 2", 1);
                wordList7.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 7th list").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);

                WordList wordList8 = document.AddList(WordListStyle.Bulleted);
                wordList8.AddItem("Text 8");
                wordList8.AddItem("Text 8.1", 1);
                wordList8.AddItem("Text 8.2", 2);
                wordList8.AddItem("Text 8.3", 2);
                wordList8.AddItem("Text 8.4", 0);
                wordList8.AddItem("Text 8.5", 0);
                wordList8.AddItem("Text 8.6", 1);

                Console.WriteLine("+ Paragraphs count: " + document.Paragraphs.Count);
                Console.WriteLine("+ Lists count: " + document.Lists.Count);

                document.Save(openWord);
            }
        }


    }
}
