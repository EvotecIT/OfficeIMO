using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlHyperlinks(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlHyperlinks.docx");
            string html = "<p>Visit <a href=\"https://evotec.xyz\">Evotec</a> or go to https://github.com</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
