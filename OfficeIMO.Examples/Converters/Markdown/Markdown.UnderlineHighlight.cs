using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownUnderlineHighlight(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownUnderlineHighlight.docx");
            using var doc = WordDocument.Create();

            var paragraph = doc.AddParagraph();
            paragraph.AddText("underlined").Underline = UnderlineValues.Single;
            paragraph.AddText(" and ");
            paragraph.AddText("highlighted").Highlight = HighlightColorValues.Yellow;

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions {
                EnableUnderline = true,
                EnableHighlight = true
            });
            Console.WriteLine(markdown);

            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

