using System;
using System.IO;
using OfficeIMO.Converters;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlRoundTrip(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlRoundTrip.docx");
            string html = "<p>Hello <b>world</b> and <i>universe</i>.</p>";

            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());
            ConverterRegistry.Register("word->html", () => new WordToHtmlConverter());

            using MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream wordStream = new MemoryStream();
            IWordConverter htmlToWord = ConverterRegistry.Resolve("html->word");
            htmlToWord.Convert(htmlStream, wordStream, new HtmlToWordOptions { FontFamily = "Calibri" });
            File.WriteAllBytes(filePath, wordStream.ToArray());

            wordStream.Position = 0;
            using MemoryStream htmlOutput = new MemoryStream();
            IWordConverter wordToHtml = ConverterRegistry.Resolve("word->html");
            wordToHtml.Convert(wordStream, htmlOutput, new WordToHtmlOptions { IncludeFontStyles = true });
            string roundTrip = Encoding.UTF8.GetString(htmlOutput.ToArray());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
