using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlCitations(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlCitations.docx");
            string html = "<p>This is a <cite>citation</cite> example.</p>";

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

