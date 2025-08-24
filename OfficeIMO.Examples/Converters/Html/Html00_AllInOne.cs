using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html00_AllInOne(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML All-In-One: comprehensive input (all.html)");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "all.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");

            string html = File.ReadAllText(htmlPath);
            var baseDir = Path.GetDirectoryName(htmlPath)!;
            using var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri", BasePath = baseDir });

            string docxPath = Path.Combine(folderPath, "Html00_AllInOne.docx");
            doc.Save(docxPath);

            string htmlOutPath = Path.Combine(folderPath, "Html00_AllInOne.roundtrip.html");
            File.WriteAllText(htmlOutPath, doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true, IncludeListStyles = true }));

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOutPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
