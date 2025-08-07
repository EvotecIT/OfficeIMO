using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTableStyles.docx");
            string html = "<table style=\"border:2px solid #ff0000; background-color:#00ff00\"><tr><td style=\"border:1px dashed #0000ff\">Cell</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
