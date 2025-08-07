using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlFootnotes(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlFootnotes.docx");

            using var doc = WordDocument.Create(filePath);
            doc.AddParagraph("Text with footnote").AddFootNote("Example footnote");
            doc.Save();

            string html = doc.ToHtml(new WordToHtmlOptions { ExportFootnotes = true });
            Console.WriteLine(html);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

