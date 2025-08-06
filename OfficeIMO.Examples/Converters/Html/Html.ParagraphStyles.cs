using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlParagraphStyles(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlParagraphStyles.docx");
            string html = "<p style=\"color:red;background-color:cyan;font-size:24px\">Styled paragraph</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
