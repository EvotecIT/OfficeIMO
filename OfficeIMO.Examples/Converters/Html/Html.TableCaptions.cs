using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableCaptions(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTableCaptions.docx");
            string html = "<table><caption>Sample caption</caption><tr><td>Cell</td></tr></table>";

            var options = new HtmlToWordOptions { TableCaptionPosition = TableCaptionPosition.Below };
            var doc = html.LoadFromHtml(options);

            doc.Save(filePath);
            string roundTrip = doc.ToHtml(new WordToHtmlOptions());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
