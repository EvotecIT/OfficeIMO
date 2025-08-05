using System;
using System.IO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownGenericFont(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownGenericFont.docx");
            string markdown = "Generic font sample.";

            ConverterRegistry.Register("markdown->word", () => new MarkdownToWordConverter());
            using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
            using MemoryStream output = new MemoryStream();
            IWordConverter converter = ConverterRegistry.Resolve("markdown->word");
            converter.Convert(input, output, new MarkdownToWordOptions { FontFamily = "monospace" });
            File.WriteAllBytes(filePath, output.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

