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
        internal static void Example_BasicLists2Load(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading standard document with lists");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists2.docx");

            using (WordDocument document = WordDocument.Load(filePath)) {
                // change on loaded document
                document.Lists[1].ListItems[3].ListItemLevel = 1;

                var paragraph = document.AddParagraph("This is 9th list").SetColor(Color.MediumAquamarine).SetUnderline(UnderlineValues.Double);

                WordList wordList8 = document.AddList(WordListStyle.Bulleted);
                wordList8.AddItem("Text 9");
                wordList8.AddItem("Text 9.1", 1);
                wordList8.AddItem("Text 9.2", 2);
                wordList8.AddItem("Text 9.3", 2);
                wordList8.AddItem("Text 9.4", 0);
                wordList8.AddItem("Text 9.5", 0);
                wordList8.AddItem("Text 9.6", 1);

                paragraph = document.AddParagraph("This is 10th list").SetColor(Color.ForestGreen).SetUnderline(UnderlineValues.Double);

                WordList wordList2 = document.AddList(WordListStyle.Numbered);
                wordList2.AddItem("Temp 10");
                wordList2.AddItem("Text 10.1", 1);

                paragraph = document.AddParagraph("Paragraph in the middle of the list").SetColor(Color.Aquamarine); //.SetUnderline(UnderlineValues.Double);

                wordList2.AddItem("Text 10.2", 2);
                wordList2.AddItem("Text 10.3", 2);

                paragraph = document.AddParagraph("This is 10th list").SetColor(Color.ForestGreen).SetUnderline(UnderlineValues.Double);

                WordList wordList3 = document.AddList(WordListStyle.Numbered);
                wordList3.AddItem("Temp 11");
                wordList3.AddItem("Text 11.1", 1);

                Console.WriteLine("+ Paragraphs count: " + document.Paragraphs.Count);
                Console.WriteLine("+ Lists count: " + document.Lists.Count);

                Console.WriteLine("+ List element 0 text: " + document.Lists[0].ListItems[0].Text);
                Console.WriteLine("+ List element 1 text: " + document.Lists[0].ListItems[1].Text);
                Console.WriteLine("+ List element 2 text: " + document.Lists[0].ListItems[2].Text);
                document.Save(openWord);
            }
        }

    }
}
