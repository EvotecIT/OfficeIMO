using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlSupSub(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlSupSub.docx");
            string html = "<p>H<sub>2</sub>O is water and E=mc<sup>2</sup>.</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

