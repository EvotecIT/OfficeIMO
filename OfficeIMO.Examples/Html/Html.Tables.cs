using OfficeIMO.Html;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTables(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTables.docx");
            string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td><table><tr><td>Nested</td></tr></table></td></tr></table>";

            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());
            ConverterRegistry.Register("word->html", () => new WordToHtmlConverter());

            using MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream wordStream = new MemoryStream();
            IWordConverter htmlToWord = ConverterRegistry.Resolve("html->word");
            htmlToWord.Convert(htmlStream, wordStream, new HtmlToWordOptions());
            File.WriteAllBytes(filePath, wordStream.ToArray());

            wordStream.Position = 0;
            using MemoryStream htmlOutput = new MemoryStream();
            IWordConverter wordToHtml = ConverterRegistry.Resolve("word->html");
            wordToHtml.Convert(wordStream, htmlOutput, new WordToHtmlOptions());
            string roundTrip = Encoding.UTF8.GetString(htmlOutput.ToArray());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
