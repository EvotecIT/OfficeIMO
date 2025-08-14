using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Converters.Html {
    internal static class HtmlStylesheetClasses {
        public static void Example_HtmlStylesheetClasses(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlStylesheetClasses.docx");
            string html = "<style>.title{font-weight:bold;font-size:32px;}</style><p class=\"title\">Styled text</p>";
            using var document = html.LoadFromHtml();
            document.Save(filePath);
            Console.WriteLine($"Document saved to: {filePath}");
            if (openWord) WordDocument.Open(filePath);
        }
    }
}
