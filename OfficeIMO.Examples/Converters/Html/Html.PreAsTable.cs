using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlPreAsTable(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlPreAsTable.docx");
            string html = "<pre><code>Console.WriteLine(\"Hello\");\nConsole.WriteLine(\"World\");</code></pre>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions { RenderPreAsTable = true });
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
