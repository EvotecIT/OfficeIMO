using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;
using System.Threading.Tasks;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html07_Async {
        public static async Task Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating async HTML conversion");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Async HTML");
            await doc.AddHtmlToHeaderAsync("<p>Header async</p>");
            await doc.AddHtmlToFooterAsync("<p>Footer async</p>");

            string outputPath = Path.Combine(folderPath, "HtmlAsync.html");
            await doc.SaveAsHtmlAsync(outputPath);

            string html = await doc.ToHtmlAsync();
            using var roundTrip = await html.LoadFromHtmlAsync();

            Console.WriteLine($"✓ Created: {outputPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}
