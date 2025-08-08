using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableSections(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlTableSections.docx");
            string html = "<table><colgroup><col style=\"width:20%\"><col style=\"width:80%\"></colgroup><thead><tr style=\"background-color:#ff0000\"><th>Header1</th><th>Header2</th></tr></thead><tbody><tr><td>Body1</td><td>Body2</td></tr></tbody><tfoot><tr><td>Foot1</td><td>Foot2</td></tr></tfoot></table>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
