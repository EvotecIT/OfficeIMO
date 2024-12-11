using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_BasicLists4(string folderPath, bool openWord) {
            using var document = WordDocument.Create();
            var listOfListStyles = (WordListStyle[])Enum.GetValues(typeof(WordListStyle));
            foreach (var listStyle in listOfListStyles) {
                var paragraph = document.AddParagraph(listStyle.ToString());
                paragraph.SetColor(Color.Red).SetBold();
                paragraph.ParagraphAlignment = JustificationValues.Center;

                var wordList1 = document.AddList(listStyle);
                if (listStyle == WordListStyle.Chapters) {
                    // chapters supports only 0 level in lists
                    wordList1.AddItem("Text 1");
                    wordList1.AddItem("Text 2");
                    wordList1.AddItem("Text 3");
                    wordList1.AddItem("Text 4");
                    wordList1.AddItem("Text 5");
                    wordList1.AddItem("Text 6");
                    wordList1.AddItem("Text 7");
                    wordList1.AddItem("Text 8");
                    wordList1.AddItem("Text 9");
                    wordList1.AddItem("Text 10");
                } else {
                    // all other lists have up to 9 level
                    wordList1.AddItem("Text 1", 0);
                    wordList1.AddItem("Text 2", 1);
                    wordList1.AddItem("Text 3", 2);
                    wordList1.AddItem("Text 4", 3);
                    wordList1.AddItem("Text 5", 4);
                    wordList1.AddItem("Text 6", 5);
                    wordList1.AddItem("Text 7", 6);
                    wordList1.AddItem("Text 8", 7);
                    wordList1.AddItem("Text 9", 8);
                }
            }

            Console.WriteLine("+ Lists Count: " + document.Lists.Count);
            Console.WriteLine("+ Lists Count: " + document.Sections[0].Lists.Count);

            var filePath = Path.Combine(folderPath, "Document with Lists4.docx");
            using var outputStream = new MemoryStream();
            document.Save(outputStream);
            File.WriteAllBytes(filePath, outputStream.ToArray());

            Helpers.Open(filePath, openWord);
        }
    }
}
