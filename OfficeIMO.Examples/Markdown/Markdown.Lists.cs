using System;
using System.IO;
using OfficeIMO.Converters;
using OfficeIMO.Markdown;

namespace OfficeIMO.Examples.Markdown {
    internal static partial class Markdown {
        public static void Example_MarkdownLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "MarkdownLists.docx");
            string markdown = "- Item 1\n- Item 2\n\n1. First\n1. Second";

            ConverterRegistry.Register("markdown->word", () => new MarkdownToWordConverter());
            ConverterRegistry.Register("word->markdown", () => new WordToMarkdownConverter());

            using MemoryStream markdownStream = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
            using MemoryStream wordStream = new MemoryStream();
            IWordConverter mdToWord = ConverterRegistry.Resolve("markdown->word");
            mdToWord.Convert(markdownStream, wordStream, new MarkdownToWordOptions());
            File.WriteAllBytes(filePath, wordStream.ToArray());

            wordStream.Position = 0;
            using MemoryStream markdownOutput = new MemoryStream();
            IWordConverter wordToMd = ConverterRegistry.Resolve("word->markdown");
            wordToMd.Convert(wordStream, markdownOutput, new WordToMarkdownOptions());
            string roundTrip = Encoding.UTF8.GetString(markdownOutput.ToArray());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
