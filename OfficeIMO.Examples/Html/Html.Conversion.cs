using System;
using System.IO;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlRoundTrip(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlRoundTrip.docx");
            string html = "<p>Hello <b>world</b> and <i>universe</i>.</p>";

            using (MemoryStream ms = new MemoryStream()) {
                HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions { FontFamily = "Calibri" });
                File.WriteAllBytes(filePath, ms.ToArray());

                ms.Position = 0;
                string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions { IncludeStyles = true });
                Console.WriteLine(roundTrip);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
