using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlComplexStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlComplexStyles.docx");
            string html = "<p style=\"margin:10pt 20pt;line-height:1.5;background-color:#ffff00\">Complex <span style=\"text-decoration:underline line-through;background-color:#00ff00\">styled</span> paragraph</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
