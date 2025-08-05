using System;
using System.IO;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownLists.docx");
            string markdown = "- Item 1\n- Item 2\n\n1. First\n1. Second";

            // Convert Markdown to Word document
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            
            // Save the Word document
            doc.Save(filePath);
            
            // Convert back to Markdown
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}