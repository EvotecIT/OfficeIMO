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
            document.AddParagraph("Paragraph three").AddFootNote("Third footnote");

            document.Save(filePath);
            string markdown = document.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(markdown);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        public static void Example_MarkdownFootNotes_Load(string folderPath, bool openWord) {
            string markdown = "Text with first[^1] and second[^2] notes.\n\nAnother line[^3].\n\n[^1]: First footnote\n[^2]: Second footnote\n[^3]: Third footnote\n";
            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            string filePath = Path.Combine(folderPath, "MarkdownFootNotesFromMarkdown.docx");
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

