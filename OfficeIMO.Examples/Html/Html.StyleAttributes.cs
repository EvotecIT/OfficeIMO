using System;
using System.IO;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlStyleAttributes(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStyleAttributes.docx");
            string html = "<p style=\"font-weight:bold;font-size:32px\">Heading 1</p>";

            ConverterRegistry.Register("html->word", () => new HtmlToWordConverter());

            using MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html));
            using MemoryStream wordStream = new MemoryStream();
            IWordConverter htmlToWord = ConverterRegistry.Resolve("html->word");
            htmlToWord.Convert(htmlStream, wordStream, new HtmlToWordOptions());
            File.WriteAllBytes(filePath, wordStream.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
