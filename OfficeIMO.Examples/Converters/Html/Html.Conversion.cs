using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlRoundTrip(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlRoundTrip.docx");
            string html = "<p>Hello <b>world</b> and <i>universe</i>.</p>";

            // Convert HTML to Word document
            var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });
            
            // Save the Word document
            doc.Save(filePath);
            
            // Convert back to HTML
            string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true });
            Console.WriteLine(roundTrip);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
        
        // TODO: Add more example methods as needed
        public static void Example_HtmlInterface(string folderPath, bool openWord) {
            // Placeholder for HTML interface example
            Example_HtmlRoundTrip(folderPath, openWord);
        }
    }
}
