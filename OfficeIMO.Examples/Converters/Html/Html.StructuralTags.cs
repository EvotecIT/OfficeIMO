using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlStructuralTags(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStructuralTags.docx");
            string html = "<address style=\"text-align:right\"><p>Location</p></address>" +
                          "<article style=\"text-align:center\"><p>Article text</p></article>" +
                          "<aside style=\"text-align:justify\"><p>Aside note</p></aside>" +
                          "<nav style=\"margin-left:20pt;padding-left:10pt\"><p>Menu</p></nav>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

