using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownFootNotes(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownFootNotes.docx");

            using var document = WordDocument.Create();
            document.AddParagraph("Paragraph one").AddFootNote("First footnote");
            document.AddParagraph("Paragraph two").AddFootNote("Second footnote");

            document.Save(filePath);
            string markdown = document.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(markdown);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

