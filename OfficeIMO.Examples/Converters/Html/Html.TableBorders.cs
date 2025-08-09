using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableBorders(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTableBorders.docx");
            string html = "<table border=\"2\"><tr><td>A1</td><td style=\"border:1px solid #ff0000\">B1</td></tr></table>";
            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

