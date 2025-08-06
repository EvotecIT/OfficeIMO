using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html03_SaveAsHtmlWithImagesAndMetadata {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document with image and metadata and saving as HTML");

            using var doc = WordDocument.Create();

            doc.BuiltinDocumentProperties.Title = "Sample HTML";
            doc.BuiltinDocumentProperties.Creator = "OfficeIMO";

            doc.AddParagraph("Example Document").Style = WordParagraphStyles.Heading1;

            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Assets", "OfficeIMO.png");
            doc.AddParagraph().AddImage(assetPath);

            string outputPath = Path.Combine(folderPath, "SaveAsHtmlWithImagesAndMetadata.html");
            doc.SaveAsHtml(outputPath, new WordToHtmlOptions { IncludeFontStyles = true, IncludeListStyles = true });

            Console.WriteLine($"✓ Created: {outputPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}

