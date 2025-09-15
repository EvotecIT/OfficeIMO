using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static async Task Example_Html10_OptionsAndAsync(string folderPath, bool openWord) {
            Console.WriteLine("[*] HTML Options + Async: head tags, classes, async round-trip");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Async and Options").Style = WordParagraphStyles.Heading1;
            var p = doc.AddParagraph("Styled text ");
            p.AddText("bold").Bold = true;

            // Save HTML synchronously with basic options
            string syncPath = Path.Combine(folderPath, "Html10_Options.sync.html");
            var options = new WordToHtmlOptions {
                IncludeFontStyles = true,
                IncludeListStyles = true
            };
            options.AdditionalMetaTags.Add(("x-example", "demo"));
            // Example link tag
            options.AdditionalLinkTags.Add(("stylesheet", "/css/site.css"));
            options.IncludeDefaultCss = true;
            doc.SaveAsHtml(syncPath, options);

            // Also generate HTML asynchronously
            string asyncPath = Path.Combine(folderPath, "Html10_Options.async.html");
            await doc.SaveAsHtmlAsync(asyncPath);

            // Round-trip both sync and async html strings
            string htmlSync = doc.ToHtml(new WordToHtmlOptions { IncludeFontStyles = true, IncludeDefaultCss = true });
            using var roundTripSync = htmlSync.LoadFromHtml(new HtmlToWordOptions { FontFamily = "Calibri" });

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
