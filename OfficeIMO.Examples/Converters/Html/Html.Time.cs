using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTime(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTime.docx");
            string html = "<p>On <time datetime=\"2023-01-01\">2023-01-01</time> we met.</p>";

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

