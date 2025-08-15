using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTextTransformLetterSpacing(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTextTransformLetterSpacing.docx");
            string html = "<p style=\"letter-spacing:2pt;text-transform:uppercase\">Hello World</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
