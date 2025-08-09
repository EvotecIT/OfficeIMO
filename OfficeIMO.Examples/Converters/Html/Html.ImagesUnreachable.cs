using System;
using System.IO;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlImagesUnreachable(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlImageUnreachable.docx");
            string html = "<p><img src=\"http://localhost:1/missing.png\" alt=\"Missing\" /></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);
            Console.WriteLine($"Images: {doc.Images.Count}, Paragraphs: {doc.Paragraphs.Count}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
