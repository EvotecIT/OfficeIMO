using System;
using System.IO;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownImages(string folderPath, bool openWord) {
            string assets = Path.Combine(AppContext.BaseDirectory, "..", "Assets");
            string localImage = Path.Combine(assets, "OfficeIMO.png");
            string markdown = $"![Local image]({localImage} \"Local description\" =100x100)\n" +
                               "![Remote image](https://via.placeholder.com/120 \"Remote description\" =120x80)\n" +
                               $"![Native size]({localImage} \"No size hints\")";

            var doc = markdown.LoadFromMarkdown(new MarkdownToWordOptions());
            string filePath = Path.Combine(folderPath, "MarkdownImages.docx");
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
