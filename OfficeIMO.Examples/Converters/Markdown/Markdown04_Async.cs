using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown04_Async {
        public static async Task Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document and saving/loading Markdown asynchronously");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Async paragraph");

            string markdownPath = Path.Combine(folderPath, "AsyncMarkdown.md");
            await doc.SaveAsMarkdownAsync(markdownPath);

            using var loaded = await markdownPath.LoadFromMarkdownAsync();
            Console.WriteLine($"Loaded paragraphs: {loaded.Paragraphs.Count}");

            Console.WriteLine($"✓ Created: {markdownPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(markdownPath) { UseShellExecute = true });
            }
        }
    }
}

