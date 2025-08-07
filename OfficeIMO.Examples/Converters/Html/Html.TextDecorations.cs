using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTextDecorations(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTextDecorations.docx");
            string html = "<p><s>strike</s> <del>delete</del> <ins>insert</ins> <mark>mark</mark></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
