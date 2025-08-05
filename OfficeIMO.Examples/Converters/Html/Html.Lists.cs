using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlLists(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlLists.docx");
            string html = "<ul><li>Item 1<ul><li>Sub 1</li><li>Sub 2</li></ul></li><li>Item 2</li></ul><ol><li>First</li><li>Second</li></ol>";

            // Convert HTML to Word document
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            
            // Save the Word document
            doc.Save(filePath);
            
            // Convert back to HTML
            string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}