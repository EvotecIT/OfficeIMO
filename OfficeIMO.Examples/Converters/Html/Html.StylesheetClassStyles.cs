using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlStylesheetClassStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStylesheetClassStyles.docx");
            string html = "<style>.title { font-weight:bold; font-size:32px; }</style><p class=\"title\">Title</p><p>Content</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

