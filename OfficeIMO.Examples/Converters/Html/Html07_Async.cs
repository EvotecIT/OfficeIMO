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

            string syncPath = Path.Combine(folderPath, "HtmlSync.html");
            doc.SaveAsHtml(syncPath);

            string asyncPath = Path.Combine(folderPath, "HtmlAsync.html");
            await doc.SaveAsHtmlAsync(asyncPath);

            string htmlSync = doc.ToHtml();
            using var roundTripSync = htmlSync.LoadFromHtml();

            string htmlAsync = await doc.ToHtmlAsync();
            using var roundTripAsync = await htmlAsync.LoadFromHtmlAsync();

            Console.WriteLine($"✓ Created: {syncPath}");
            Console.WriteLine($"✓ Created: {asyncPath}");

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(syncPath) { UseShellExecute = true });
            }
        }
    }
}
