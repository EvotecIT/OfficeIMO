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
        internal static void Example_BasicListsWithChangedStyling(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists with custom styling.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("This is 1st list - LowerLetterWithBracket");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(WordListStyle.LowerLetterWithBracket);
                wordList1.Bold = true;
                wordList1.FontSize = 16;
                wordList1.Color = Color.DarkRed;

                var listItem1 = wordList1.AddItem("Text 1");
                listItem1.Bold = true;
                listItem1.FontSize = 16;
                listItem1.Color = Color.DarkRed;

                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 2nd list - LowerLetterWithDot");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.FontSize = 16;
                paragraph.Color = Color.AliceBlue;

                document.Save(openWord);
            }
        }
    }
}
