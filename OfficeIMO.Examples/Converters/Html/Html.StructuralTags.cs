using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlStructuralTags(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStructuralTags.docx");
            string html = "<article id=\"art1\" style=\"text-align:center\"><p>Article text</p></article>" +
                          "<aside id=\"note1\" style=\"text-align:justify\"><p>Aside note</p></aside>" +
                          "<nav id=\"menu1\" style=\"margin-left:20pt;padding-left:10pt\"><p>Menu</p></nav>";

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            System.Console.WriteLine(string.Join(", ", doc.Bookmarks.Select(b => b.Name)));
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}

