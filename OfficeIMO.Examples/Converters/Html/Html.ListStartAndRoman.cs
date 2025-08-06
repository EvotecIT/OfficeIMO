using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_HtmlListStartAndRoman(string folderPath, bool openWord) {
            string filePath = Path.Combine(folderPath, "HtmlListStartAndRoman.docx");

            using var doc = WordDocument.Create();
            var list = doc.AddList(WordListStyle.Headings111);
            list.Numbering.Levels[0].SetStartNumberingValue(3);
            list.AddItem("Third");
            list.AddItem("Fourth");

            var roman = doc.AddList(WordListStyle.HeadingIA1);
            roman.AddItem("Intro");
            roman.AddItem("Body");

            string html = doc.ToHtml(new WordToHtmlOptions { IncludeListStyles = true });
            Console.WriteLine(html);

            doc.Save(filePath);
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
