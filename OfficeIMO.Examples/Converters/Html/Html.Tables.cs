using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTables(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTables.docx");
            string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td><table><tr><td>Nested</td></tr></table></td></tr></table>";

            // Convert HTML to Word document
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            
            // Save the Word document
            doc.Save(filePath);
            
            // Convert back to HTML
            string roundTrip = doc.ToHtml(new WordToHtmlOptions());
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}