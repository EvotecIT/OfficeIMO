using System;
using System.IO;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownLists.docx");
            string markdown = "- Item 1\n- Item 2\n\n1. First\n1. Second";

            using (MemoryStream ms = new MemoryStream()) {
                MarkdownToWordConverter.Convert(markdown, ms, new MarkdownToWordOptions());
                File.WriteAllBytes(filePath, ms.ToArray());

                ms.Position = 0;
                string roundTrip = WordToMarkdownConverter.Convert(ms, new WordToMarkdownOptions());
                Console.WriteLine(roundTrip);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
