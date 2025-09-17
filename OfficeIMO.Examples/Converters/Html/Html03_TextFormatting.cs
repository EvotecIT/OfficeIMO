using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html03_TextFormatting(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Text Formatting: spans, decorations, sup/sub");

            string html = @"<p>
  <span style='text-decoration: underline'>underlined</span>,
  <span style='text-decoration: line-through'>strikethrough</span>,
  <sup>superscript</sup>, <sub>subscript</sub>,
  <span style='text-transform: uppercase'>upper</span>,
  <span style='letter-spacing: 2px'>spaced</span>
</p>";

            using var doc = html.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });
            string docxPath = Path.Combine(folderPath, "Html03_TextFormatting.docx");
            doc.Save(docxPath);

            string htmlOut = Path.Combine(folderPath, "Html03_TextFormatting.roundtrip.html");
            File.WriteAllText(htmlOut, doc.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true, IncludeFontStyles = true }));

            Console.WriteLine($"✓ Created: {docxPath}");
            Console.WriteLine($"✓ Round-trip HTML: {htmlOut}");
            Console.WriteLine($"Paragraphs: {doc.Paragraphs.Count}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
            }
        }
    }
}

