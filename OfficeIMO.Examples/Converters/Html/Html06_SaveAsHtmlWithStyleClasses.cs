using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html06_SaveAsHtmlWithStyleClasses {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document with styles exported as classes");

            using var doc = WordDocument.Create();
            var paragraph = doc.AddParagraph("Sample Heading");
            paragraph.Style = WordParagraphStyles.Heading1;
            paragraph.AddText(" with styled run").CharacterStyleId = "Heading1Char";

            string outputPath = Path.Combine(folderPath, "SaveAsHtmlWithStyleClasses.html");
            var options = new WordToHtmlOptions { IncludeParagraphClasses = true, IncludeRunClasses = true };
            doc.SaveAsHtml(outputPath, options);

            Console.WriteLine($"✓ Created: {outputPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}
