using System;
using System.IO;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownRoundTrip(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownRoundTrip.docx");
            string markdown = "# Heading 1\n\nHello **world** and *universe*.";

            // Convert Markdown to Word document
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions { FontFamily = "Calibri" });
            
            // Save the Word document
            doc.Save(filePath);
            
            // Convert back to Markdown
            string roundTrip = doc.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
        
        // TODO: Add more example methods as needed
        public static void Example_MarkdownInterface(string folderPath, bool openWord) {
            // Placeholder for Markdown interface example
            Example_MarkdownRoundTrip(folderPath, openWord);
        }
    }
}