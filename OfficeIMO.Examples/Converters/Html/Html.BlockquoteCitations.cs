using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlBlockquoteCitations(string folderPath, bool openWord) {
            string html = "<blockquote cite=\"https://example.com\">Quoted text</blockquote>";
            using WordDocument document = html.LoadFromHtml(new HtmlToWordOptions());
            Console.WriteLine("Footnotes count: " + document.FootNotes.Count);
            if (document.FootNotes.Count > 0) {
                Console.WriteLine("Citation: " + document.FootNotes[0].Paragraphs[1].Text);
            }
            string filePath = Path.Combine(folderPath, "HtmlBlockquoteCitation.docx");
            document.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}