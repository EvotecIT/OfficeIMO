using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html04_SaveAsHtmlWithImagePaths {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document with image paths and saving as HTML");

            using var doc = WordDocument.Create();

            string assetPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Assets", "OfficeIMO.png");
            doc.AddParagraph("Image Example").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph().AddImage(assetPath, 100, 100);

            string outputPath = Path.Combine(folderPath, "SaveAsHtmlWithImagePaths.html");
            doc.SaveAsHtml(outputPath, new WordToHtmlOptions { EmbedImagesAsBase64 = false });

            Console.WriteLine($"✓ Created: {outputPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}
