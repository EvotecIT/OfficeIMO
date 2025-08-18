using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Html.Converters;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html08_Converters {
        public static async Task Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating direct converter usage (sync and async)");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Hello from converter");

            var toHtml = new WordToHtmlConverter();
            string htmlSync = toHtml.Convert(doc, new WordToHtmlOptions());
            string htmlAsync = await toHtml.ConvertAsync(doc, new WordToHtmlOptions());

            var toWord = new HtmlToWordConverter();
            using var docSync = toWord.Convert(htmlSync, new HtmlToWordOptions());
            using var docAsync = await toWord.ConvertAsync(htmlAsync, new HtmlToWordOptions());

            string htmlPath = Path.Combine(folderPath, "HtmlConverterSync.html");
            File.WriteAllText(htmlPath, htmlSync);
            Console.WriteLine($"✓ Created: {htmlPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true });
            }
        }
    }
}
