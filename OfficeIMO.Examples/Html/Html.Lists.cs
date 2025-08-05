using OfficeIMO.Converters;
using OfficeIMO.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlLists.docx");
            string html = "<ul><li>Item 1<ul><li>Sub 1</li><li>Sub 2</li></ul></li><li>Item 2</li></ul><ol><li>First</li><li>Second</li></ol>";

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
            wordToHtml.Convert(wordStream, htmlOutput, new WordToHtmlOptions { IncludeListStyles = true });
            string roundTrip = Encoding.UTF8.GetString(htmlOutput.ToArray());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
