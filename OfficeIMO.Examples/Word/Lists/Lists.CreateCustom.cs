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
        internal static void Example_CustomList1(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Custom Lists 1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("This is 1st list");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(WordListStyle.Bulleted);
                wordList1.AddItem("Text 1 - List1");
                wordList1.AddItem("Text 2 - List1", 1);
                wordList1.AddItem("Text 3 - List1", 2);

                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                // remove single level
                wordList1.Numbering.Levels[0].Remove();
                // remove all levels
                wordList1.Numbering.RemoveAllLevels();


                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                var level = new WordListLevel(NumberFormatValues.Bullet);
                wordList1.Numbering.AddLevel(level);

                var level1 = new WordListLevel(NumberFormatValues.Decimal);
                wordList1.Numbering.AddLevel(level1);

                var level2 = new WordListLevel(NumberFormatValues.Bullet);
                wordList1.Numbering.AddLevel(level2);

                Console.WriteLine("Current levels count: " + wordList1.Numbering.Levels.Count);

                document.Save(openWord);
            }
        }
    }
}
