using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlQuotes(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlQuotes.docx");
            string html = "<p>Before <q>quoted</q> after</p>";

            using var document = html.LoadFromHtml();
            document.Save(filePath);

            string roundTrip = document.ToHtml();
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
