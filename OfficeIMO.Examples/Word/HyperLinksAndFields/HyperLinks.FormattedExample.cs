using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class HyperLinks {

        internal static void Example_FormattedHyperLinks(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with formatted hyperlinks");
            string filePath = Path.Combine(folderPath, "FormattedHyperLinks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Search using ");
                var google = paragraph.AddHyperLink("Google", new Uri("https://google.com"), addStyle: true);
                google.Bold = true;
                var reference = google.Hyperlink;

                reference.InsertFormattedHyperlinkAfter("Bing", new Uri("https://bing.com"));

                document.Save(openWord);
            }
        }
    }
}

