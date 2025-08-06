using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlListStartAndType(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlListStartAndType.docx");
            string html = "<ol start=\"3\" type=\"I\"><li>Third</li><li>Fourth</li></ol><ul type=\"square\"><li>Square</li></ul>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            string roundTrip = doc.ToHtml(new WordToHtmlOptions());
            Console.WriteLine(roundTrip);

            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
