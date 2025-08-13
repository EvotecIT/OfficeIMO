using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlHyperlinks(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlHyperlinks.docx");
            string html = "<p id=\"top\">Top</p><p>Visit <a href=\"https://evotec.xyz\" title=\"Evotec site\" target=\"_blank\">Evotec</a> or <a href=\"#top\" title=\"Back to top\">back to top</a></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions { SupportsAnchorLinks = true });
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
