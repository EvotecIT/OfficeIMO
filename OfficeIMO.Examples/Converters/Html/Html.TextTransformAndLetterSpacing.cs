using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTextTransformAndLetterSpacing(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTextTransformAndLetterSpacing.docx");
            string html = "<p style=\"text-transform:uppercase;letter-spacing:2pt\">Example Text</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
