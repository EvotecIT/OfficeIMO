using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html06_ImagesAllModes(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Images: data URI, relative, absolute");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "v1-images.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");
            string html = File.ReadAllText(htmlPath);

            var baseDir = Path.GetDirectoryName(htmlPath)!;
            using var doc = html.LoadFromHtml(new HtmlToWordOptions { BasePath = baseDir });
            string docxPath = Path.Combine(folderPath, "Html06_ImagesAllModes.docx");
            try {
                doc.Save(docxPath);
            } catch (Exception ex) {
                Console.WriteLine($"! Skipping save due to: {ex.Message}");
            }

            string htmlOut = Path.Combine(folderPath, "Html06_ImagesAllModes.roundtrip.html");
            File.WriteAllText(htmlOut, doc.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true }));

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
