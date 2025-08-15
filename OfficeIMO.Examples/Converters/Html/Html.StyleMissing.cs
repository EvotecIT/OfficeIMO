using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlStyleMissing(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStyleMissing.docx");
            string html = "<p class=\"warning\">Warning</p>";
            EventHandler<StyleMissingEventArgs> handler = (s, e) => {
                if (e.ClassName == "warning") {
                    WordParagraphStyle.RegisterFontStyle("WarningStyle", "Courier New");
                    e.StyleId = "WarningStyle";
                }
            };
            WordHtmlConverterExtensions.StyleMissing += handler;
            var doc = html.LoadFromHtml();
            WordHtmlConverterExtensions.StyleMissing -= handler;
            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
