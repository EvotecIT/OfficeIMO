using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlWhiteSpace(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlWhiteSpace.docx");
            string html = "<p style=\"white-space:normal\">Hello   world\nFoo</p>" +
                          "<p style=\"white-space:pre\">Hello   world\nFoo</p>" +
                          "<p style=\"white-space:pre-wrap\">Hello   world\nFoo</p>" +
                          "<p style=\"white-space:nowrap\">Hello   world\nFoo</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            doc.Save(filePath);

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
