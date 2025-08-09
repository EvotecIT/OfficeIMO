using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlNestedTables(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlNestedTables.docx");
            string html = "<table><tr><td>Outer</td><td><table><tr><td>Inner</td></tr></table></td></tr></table>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            doc.Save(filePath);

            var outer = doc.Sections[0].Tables[0];
            var inner = outer.Rows[0].Cells[1].NestedTables[0];
            Console.WriteLine("Outer cell text: " + outer.Rows[0].Cells[0].Paragraphs[0].Text);
            Console.WriteLine("Inner cell text: " + inner.Rows[0].Cells[0].Paragraphs[0].Text);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
