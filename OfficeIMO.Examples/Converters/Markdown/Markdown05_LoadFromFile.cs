using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown05_LoadFromFile {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Loading Markdown from file");

            string mdPath = Path.Combine(folderPath, "Sample.md");
            File.WriteAllText(mdPath, "# Title\n\nLoaded from file");

            using var document = WordMarkdownConverterExtensions.LoadFromMarkdown(mdPath, encoding: null);
            string outputPath = Path.Combine(folderPath, "LoadFromFile.docx");
            document.Save(outputPath);

            Console.WriteLine($"\u2713 Created: {outputPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}

