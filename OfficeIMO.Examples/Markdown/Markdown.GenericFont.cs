using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownGenericFont(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownGenericFont.docx");
            string markdown = "Generic font sample.";

            using MemoryStream ms = new MemoryStream();
            MarkdownToWordConverter.Convert(markdown, ms, new MarkdownToWordOptions { FontFamily = "monospace" });
            File.WriteAllBytes(filePath, ms.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

