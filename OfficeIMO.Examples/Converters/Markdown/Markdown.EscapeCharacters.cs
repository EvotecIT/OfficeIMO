using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownEscapeCharacters(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownEscapeCharacters.docx");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Characters: * _ [ ] ( ) # + - . ! \\ >");
            doc.Save(filePath);

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(markdown);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

