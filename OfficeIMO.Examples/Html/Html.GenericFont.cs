using System;
using System.IO;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlGenericFont(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlGenericFont.docx");
            string html = "<p>Generic font sample.</p>";

            using MemoryStream ms = new MemoryStream();
            HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions { FontFamily = "monospace" });
            File.WriteAllBytes(filePath, ms.ToArray());

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

