using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlClassStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlClassStyles.docx");
            string html = "<p class=\"title\">Title</p><p>Content</p>";

            var options = new HtmlToWordOptions();
            options.ClassStyles["title"] = WordParagraphStyles.Heading1;
            var doc = html.LoadFromHtml(options);
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
