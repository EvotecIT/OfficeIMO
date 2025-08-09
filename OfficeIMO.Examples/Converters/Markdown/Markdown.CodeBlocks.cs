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

        public static void Example_WordToMarkdownCodeBlocks(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "WordToMarkdownCodeBlock.docx");
            using var doc = WordDocument.Create();
            string mono = FontResolver.Resolve("monospace") ?? "Consolas";
            doc.AddParagraph("Console.WriteLine(\"Hello\");").SetFontFamily(mono).SetStyleId("CodeLang_csharp");

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());
            Console.WriteLine(markdown);

            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }

        public static void Example_WordToMarkdownCodeBlocks_CustomFont(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "WordToMarkdownCodeBlockCustomFont.docx");
            using var doc = WordDocument.Create();
            const string codeFont = "Courier New";
            doc.AddParagraph("System.out.println(\"Hello\");").SetFontFamily(codeFont).SetStyleId("CodeLang_java");

            var options = new WordToMarkdownOptions { FontFamily = codeFont };
            string markdown = doc.ToMarkdown(options);
            Console.WriteLine(markdown);

            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
