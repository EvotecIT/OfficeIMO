using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html01_LoadAndRoundTripBasics(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Basics: load → docx → round-trip html");

            string htmlPath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "v1-basic.html");
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"Missing test input: {htmlPath}");
            string html = File.ReadAllText(htmlPath);

            using var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });
            string docxPath = Path.Combine(folderPath, "Html01_Basics.docx");
            doc.Save(docxPath);

            string roundTrip = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true, IncludeListStyles = true });
            string htmlOut = Path.Combine(folderPath, "Html01_Basics.roundtrip.html");
            File.WriteAllText(htmlOut, roundTrip);

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");
            Console.WriteLine($"Paragraphs: {doc.Paragraphs.Count}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}
