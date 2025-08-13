using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlAbbreviations(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlAbbreviations.docx");
            string html = "<abbr title=\"World Health Organization\">WHO</abbr>";

            using var document = html.LoadFromHtml();
            document.Save(filePath);

            string roundTrip = document.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
