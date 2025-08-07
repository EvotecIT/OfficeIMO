using System;
using System.IO;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownCodeBlocks(string folderPath, bool openWord) {
            string markdown = "```csharp\nConsole.WriteLine(\"Hello\");\n```";

            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            var codeParagraph = doc.Paragraphs[0];
            Console.WriteLine($"Detected language style: {codeParagraph.StyleId}");

            string filePath = Path.Combine(folderPath, "MarkdownCodeBlock.docx");
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
