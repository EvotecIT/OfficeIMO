using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlComplexSelectors(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlComplexSelectors.docx");
            string html = "<style>div p.important { color:#ff0000; }</style><div><p class=\"important\">Styled</p></div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
