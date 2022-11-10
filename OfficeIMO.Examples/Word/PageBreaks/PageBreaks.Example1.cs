using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageBreaks {
        internal static void Example_PageBreaks(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with page breaks and removing them");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some page breaks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "This is my title";
                document.BuiltinDocumentProperties.Creator = "Przemysław Kłys";
                document.BuiltinDocumentProperties.Keywords = "word, docx, test";

                var paragraph = document.AddParagraph("Test 1");

                document.AddPageBreak();

                paragraph.Text = "Test 2";

                paragraph = document.AddParagraph("Test 2");

                // Now lets remove paragraph with page break
                document.Paragraphs[1].Remove();

                // Now lets remove 1st paragraph
                document.Paragraphs[0].Remove();

                document.AddPageBreak();

                document.AddParagraph().Text = "Some text on next page";

                var paragraph1 = document.AddParagraph("Test").AddText("Test2");
                paragraph1.Color = SixLabors.ImageSharp.Color.Red;
                paragraph1.AddText("Test3");

                paragraph = document.AddParagraph("Some paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AddText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                paragraph = document.AddParagraph("2nd paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AddText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AddText(" More text");
                paragraph.Color = SixLabors.ImageSharp.Color.CornflowerBlue;

                // remove last paragraph
                document.Paragraphs.Last().Remove();

                paragraph = document.AddParagraph("2nd paragraph");
                paragraph.Bold = true;
                paragraph = paragraph.AddText(" continue?");
                paragraph.Underline = UnderlineValues.DashLong;
                paragraph = paragraph.AddText(" More text");
                paragraph.Color = SixLabors.ImageSharp.Color.CornflowerBlue;

                // remove paragraph
                int countParagraphs = document.Paragraphs.Count;
                document.Paragraphs[countParagraphs - 2].Remove();

                // remove first page break
                document.PageBreaks[0].Remove(true);

                document.Save(openWord);
            }
        }




    }
}
