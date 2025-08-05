using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlImages(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlImage.docx");
            byte[] imageBytes = File.ReadAllBytes(Path.Combine("Assets", "OfficeIMO.png"));
            string base64 = Convert.ToBase64String(imageBytes);
            string html = $"<p><img src=\"data:image/png;base64,{base64}\" /></p>";

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