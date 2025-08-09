using System;
using System.IO;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownHeadingsBoldLinks(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownHeadingsBoldLinks.docx");
            string markdown = "# Heading 1\n\nThis is **bold** text with a [link](https://example.com).";

            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
