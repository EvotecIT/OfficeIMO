using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Bookmarks {
        internal static void Example_BasicWordWithBookmarks(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with bookmarks");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithBookmarks.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test 1").AddBookmark("Start");

                var paragraph = document.AddParagraph("This is text");
                foreach (string text in new List<string>() { "text1", "text2", "text3" }) {
                    paragraph = paragraph.AddText(text);
                    paragraph.Bold = true;
                    paragraph.Italic = true;
                    paragraph.Underline = UnderlineValues.DashDotDotHeavy;
                }

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 2").AddBookmark("Middle1");

                paragraph.AddText("OK baby");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 3").AddBookmark("Middle0");

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 4").AddBookmark("EndOfDocument");

                document.Bookmarks[2].Remove();

                document.AddPageBreak();
                document.AddPageBreak();

                document.AddParagraph("Test 5");

                document.PageBreaks[7].Remove(includingParagraph: false);
                document.PageBreaks[6].Remove(true);

                Console.WriteLine(document.DocumentIsValid);
                Console.WriteLine(document.DocumentValidationErrors.Count);

                document.Save(openWord);
            }
        }


    }
}
