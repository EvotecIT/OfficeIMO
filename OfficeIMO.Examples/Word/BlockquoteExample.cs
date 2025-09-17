using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static class BlockquoteExample {
        public static void Example_BlockquoteRoundTrip(string folderPath, bool openWord) {
            Console.WriteLine("[*] Blockquote round-trip HTML <-> Word");

            string html = "<blockquote>Quoted text</blockquote>";
            using (WordDocument document = html.LoadFromHtml(new HtmlToWordOptions())) {
                string docPath = Path.Combine(folderPath, "Blockquote.docx");
                document.Save(docPath);
                Console.WriteLine($"✓ Created: {docPath}");

                string roundTrip = document.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true });
                Console.WriteLine("Round-trip HTML:");
                Console.WriteLine(roundTrip);
            }
        }
    }
}
