using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html09_CodePreWhitespace(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Code/Pre: whitespace and monospaced mapping");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "v1-code-pre.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");
            string html = File.ReadAllText(htmlPath);

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            string docxPath = Path.Combine(folderPath, "Html09_CodePreWhitespace.docx");
            doc.Save(docxPath);

            string htmlOut = Path.Combine(folderPath, "Html09_CodePreWhitespace.roundtrip.html");
            File.WriteAllText(htmlOut, doc.ToHtml());

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
