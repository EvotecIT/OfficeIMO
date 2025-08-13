using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlHeaderFooter(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlHeaderFooter.docx");

            using WordDocument document = WordDocument.Create();
            document.AddHtmlToHeader("<p>Header content</p>");
            document.AddHtmlToFooter("<p>Footer content</p>");
            document.Save(filePath);
            Console.WriteLine($"Created: {filePath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
