using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html05_SaveAsHtmlWithAdditionalHeadTags {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document with extra head elements and saving as HTML");

            using var doc = WordDocument.Create();

            doc.BuiltinDocumentProperties.Title = "Sample HTML";

            doc.AddParagraph("Example Document").Style = WordParagraphStyles.Heading1;

            var options = new WordToHtmlOptions();
            options.AdditionalMetaTags.Add(("viewport", "width=device-width, initial-scale=1"));
            options.AdditionalLinkTags.Add(("stylesheet", "styles.css"));

            string outputPath = Path.Combine(folderPath, "SaveAsHtmlWithAdditionalHeadTags.html");
            doc.SaveAsHtml(outputPath, options);

            Console.WriteLine($"✓ Created: {outputPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}

