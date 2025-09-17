using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html07_LinksAndAnchors(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Links: external + internal anchors");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "v1-links-anchors.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");
            string html = File.ReadAllText(htmlPath);

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            string docxPath = Path.Combine(folderPath, "Html07_LinksAndAnchors.docx");
            doc.Save(docxPath);

            string htmlOut = Path.Combine(folderPath, "Html07_LinksAndAnchors.roundtrip.html");
            File.WriteAllText(htmlOut, doc.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true }));

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
