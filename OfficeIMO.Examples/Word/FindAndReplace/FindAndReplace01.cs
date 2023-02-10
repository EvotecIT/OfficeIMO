using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class FindAndReplace {
        internal static void Example_FindAndReplace01(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document - Find & Replace");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document to replace text.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test Section");

                document.Paragraphs[0].AddComment("Przemysław", "PK", "This is my comment");

                document.AddParagraph("Test Section - another line");

                document.Paragraphs[1].AddComment("Przemysław", "PK", "More comments");

                document.AddParagraph("This is a text ").AddText("more text").AddText(" even longer text").AddText(" and even longer right?");

                document.AddParagraph("This is a text ").AddText("more text 1").AddText(" even longer text 1").AddText(" and even longer right?");
                // we now ensure that we add bold to complicate the search
                document.Paragraphs[9].Bold = true;
                document.Paragraphs[10].Bold = true;

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var replacedCount = document.FindAndReplace("Test Section", "Production Section");
                Console.WriteLine("Replaced: " + replacedCount);

                // should be 0 because it stretches over 2 paragraphs
                var replacedCount1 = document.FindAndReplace("This is a text more text", "Shorter text");
                Console.WriteLine("Replaced (should be 0): " + replacedCount1);

                document.CleanupDocument();

                // cleanup should merge paragraphs making it easier to find and replace text
                // this only works for same formatting though
                // may require improvement in the future to ignore formatting completely, but then it's a bit tricky which formatting to apply
                var replacedCount2 = document.FindAndReplace("This is a text more text", "Shorter text");
                Console.WriteLine("Replaced (should be 1): " + replacedCount2);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                Console.WriteLine(document.Paragraphs[0].Text == "Production Section" ? "OK" : "FAIL");

                document.Save(openWord);
            }
        }
    }
}
