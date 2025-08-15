using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlTableBorderCollapse(string folderPath, bool openWord) {
            string filePathCollapsed = Path.Combine(folderPath, "HtmlTableBorderCollapse.docx");
            string collapsedHtml = "<table style=\"border-collapse:collapse;border:2px solid #ff0000\"><tr><td>A1</td><td>B1</td></tr></table>";
            using (var doc = collapsedHtml.LoadFromHtml(new HtmlToWordOptions())) {
                doc.Save(filePathCollapsed);
            }

            string filePathSeparate = Path.Combine(folderPath, "HtmlTableBorderSeparate.docx");
            string separateHtml = "<table style=\"border-collapse:separate;border:2px solid #ff0000\"><tr><td>A1</td><td>B1</td></tr></table>";
            using (var doc = separateHtml.LoadFromHtml(new HtmlToWordOptions())) {
                doc.Save(filePathSeparate);
            }

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePathCollapsed) { UseShellExecute = true });
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePathSeparate) { UseShellExecute = true });
            }
        }
    }
}

