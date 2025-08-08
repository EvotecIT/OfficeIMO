using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlImagesRelative(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlImageRelative.docx");
            string html = "<p><img src=\"OfficeIMO.png\" alt=\"Logo\"/></p>";
            var options = new HtmlToWordOptions { BasePath = "Assets" };
            var doc = html.LoadFromHtml(options);
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
