using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Embed {
        internal static void Example_EmbedFragmentAfter(string folderPath, bool openWord) {
            Console.WriteLine("[*] Embed HTML fragment after paragraph");
            string filePath = System.IO.Path.Combine(folderPath, "EmbedFragmentAfter.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("Intro");
                document.AddParagraph("End");

                string html = "<html><body><p>Inserted</p></body></html>";
                document.AddEmbeddedFragmentAfter(p1, html);

                Console.WriteLine("Embedded: " + document.EmbeddedDocuments.Count);
                document.Save(openWord);
            }
        }
    }
}

