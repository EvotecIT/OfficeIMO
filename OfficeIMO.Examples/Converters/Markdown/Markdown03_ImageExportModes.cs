using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown03_ImageExportModes {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating image export modes");

            using var doc = WordDocument.Create();
            string imagePath = Path.Combine(AppContext.BaseDirectory, "..", "Assets", "OfficeIMO.png");
            doc.AddParagraph("Image example");
            doc.AddParagraph().AddImage(imagePath);

            // Base64 export
            string base64Path = Path.Combine(folderPath, "MarkdownImageBase64.md");
            doc.SaveAsMarkdown(base64Path, new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.Base64
            });
            Console.WriteLine($"✓ Base64 markdown: {base64Path}");

            // File export
            string fileModePath = Path.Combine(folderPath, "MarkdownImageFiles.md");
            doc.SaveAsMarkdown(fileModePath, new WordToMarkdownOptions {
                ImageExportMode = ImageExportMode.File,
                ImageDirectory = folderPath
            });
            Console.WriteLine($"✓ File markdown: {fileModePath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(base64Path) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileModePath) { UseShellExecute = true });
            }
        }
    }
}
