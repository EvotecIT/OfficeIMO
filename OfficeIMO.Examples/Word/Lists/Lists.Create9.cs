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
        internal static void Example_BasicLists9(string folderPath, bool openWord) {
            string filePath = System.IO.Path.Combine(folderPath, "Document with Lists with letters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("This is 1st list - LowerLetterWithBracket");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList1 = document.AddList(WordListStyle.LowerLetterWithBracket);
                wordList1.AddItem("Text 1");
                wordList1.AddItem("Text 2", 1);
                wordList1.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 2nd list - LowerLetterWithDot");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList2 = document.AddList(WordListStyle.LowerLetterWithDot);
                wordList2.AddItem("Text 1");
                wordList2.AddItem("Text 2", 1);
                wordList2.AddItem("Text 3", 2);

                paragraph = document.AddParagraph("This is 3rd list - UpperLetterWithDot").SetColor(Color.DeepPink).SetUnderline(UnderlineValues.Double);
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList3 = document.AddList(WordListStyle.UpperLetterWithDot);
                wordList3.AddItem("Text 3.1", 1);
                wordList3.AddItem("Text 3.2", 2);
                wordList3.AddItem("Text 3.3", 2);
                wordList3.AddItem("Text 3.4", 0);
                wordList3.AddItem("Text 3.5", 0);
                wordList3.AddItem("Text 3.6", 1);
                wordList3.AddItem("Text 3");

                paragraph = document.AddParagraph("This is 4th list - UpperLetterWithBracket").SetColor(Color.Aqua).SetUnderline(UnderlineValues.Double);
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordList wordList4 = document.AddList(WordListStyle.UpperLetterWithBracket);
                wordList4.AddItem("Text 8");
                wordList4.AddItem("Text 8.1", 1);
                wordList4.AddItem("Text 8.2", 2);
                wordList4.AddItem("Text 8.3", 2);
                wordList4.AddItem("Text 8.4", 0);
                wordList4.AddItem("Text 8.5", 0);
                wordList4.AddItem("Text 8.6", 1);

                document.Save(openWord);
            }
        }
    }
}
