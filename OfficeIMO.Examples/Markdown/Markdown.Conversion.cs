using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownRoundTrip(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownRoundTrip.docx");
            string markdown = "# Heading 1\n\nHello **world** and *universe*.";

            using (MemoryStream ms = new MemoryStream()) {
                MarkdownToWordConverter.Convert(markdown, ms, new MarkdownToWordOptions { FontFamily = "Calibri" });
                File.WriteAllBytes(filePath, ms.ToArray());

                ms.Position = 0;
                string roundTrip = WordToMarkdownConverter.Convert(ms, new WordToMarkdownOptions { FontFamily = "Calibri" });
                Console.WriteLine(roundTrip);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
