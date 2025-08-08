using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlSpanStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlSpanStyles.docx");
            string html = "<p>Span with <span style=\"color:#ff0000;font-family:Arial;font-size:24px;font-weight:bold;font-style:italic\">styled text</span> <span style=\"vertical-align:super\">super</span><span style=\"vertical-align:sub\">sub</span></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
