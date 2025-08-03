using System;
using System.IO;
using OfficeIMO.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlHeadings(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlHeadings.docx");
            string html = "<h1>Heading 1</h1><h2>Heading 2</h2><h3>Heading 3</h3><h4>Heading 4</h4><h5>Heading 5</h5><h6>Heading 6</h6>";

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
