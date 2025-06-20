using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class FindAndReplace {
        internal static void Example_ReplaceTextWithHtmlFragment(string folderPath, bool openWord) {
            Console.WriteLine("[*] Replace text with HTML fragment");
            string filePath = System.IO.Path.Combine(folderPath, "ReplaceTextWithHtmlFragment.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro start");
                document.AddParagraph("finish end");

                string html = "<html><body><p>Injected via AltChunk</p></body></html>";
                int replaced = document.ReplaceTextWithHtmlFragment("startfinish", html);

                Console.WriteLine($"Replaced: {replaced}");
                Console.WriteLine("Embedded documents: " + document.EmbeddedDocuments.Count);
                document.Save(openWord);
            }
        }
    }
}
