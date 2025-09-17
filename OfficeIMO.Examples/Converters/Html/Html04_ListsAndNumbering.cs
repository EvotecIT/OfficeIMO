using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html04_ListsAndNumbering(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Lists: nested, start, roman/alpha");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "lists-deep.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");
            string html = File.ReadAllText(htmlPath);

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            string docxPath = Path.Combine(folderPath, "Html04_ListsAndNumbering.docx");
            doc.Save(docxPath);

            string htmlOut = Path.Combine(folderPath, "Html04_ListsAndNumbering.roundtrip.html");
            File.WriteAllText(htmlOut, doc.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true, IncludeListStyles = true }));

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
