using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownPageSettings(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownPageSettings.docx");
            string markdown = "Hello World";

            // Convert Markdown to Word document with page settings
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions {
                DefaultOrientation = PageOrientationValues.Landscape,
                DefaultPageSize = WordPageSize.A5
            });
            
            // Save the Word document
            doc.Save(filePath);
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}