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
        internal static void Example_BasicLists4(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists4.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var listOfListStyles = (WordListStyle[])Enum.GetValues(typeof(WordListStyle));
                foreach (var listStyle in listOfListStyles) {
                    var paragraph = document.AddParagraph(listStyle.ToString());
                    paragraph.SetColor(Color.Red).SetBold();
                    paragraph.ParagraphAlignment = JustificationValues.Center;

                    if (listStyle == WordListStyle.Chapters) {
                        // chapters supports only 0 level in lists
                        WordList wordList1 = document.AddList(listStyle);
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
                        WordList wordList1 = document.AddList(listStyle);
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
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                // change on loaded document
                document.Lists[0].ListItems[3].ListItemLevel = 1;


                document.Save(openWord);
            }
        }
    }
}
