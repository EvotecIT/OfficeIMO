using OfficeIMO.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTables(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTables.docx");
            string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td><table><tr><td>Nested</td></tr></table></td></tr></table>";

            using (MemoryStream ms = new MemoryStream()) {
                HtmlToWordConverter.Convert(html, ms, new HtmlToWordOptions());
                File.WriteAllBytes(filePath, ms.ToArray());

                ms.Position = 0;
                string roundTrip = WordToHtmlConverter.Convert(ms, new WordToHtmlOptions());
                Console.WriteLine(roundTrip);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
