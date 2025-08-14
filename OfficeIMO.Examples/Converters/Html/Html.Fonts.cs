using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlFonts(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlFonts.docx");
            string html = "<p style=\"font-family: 'Courier New', monospace\">Sample text</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
