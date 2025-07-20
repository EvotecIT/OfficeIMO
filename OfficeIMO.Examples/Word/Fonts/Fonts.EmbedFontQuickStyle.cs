using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fonts {
        public static void Example_EmbedFontWithStyle(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with embedded font and auto style");
            string filePath = Path.Combine(folderPath, "DocumentEmbeddedFontAutoStyle.docx");
            string fontPath = Path.Combine(templatesPath, "DejaVuSans.ttf");
            if (!File.Exists(fontPath)) {
                Console.WriteLine($"Font file not found: {fontPath}");
                return;
            }

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.EmbedFont(fontPath, "DejaVuStyle", "DejaVu Style");

                document.AddParagraph("Paragraph using registered style").SetStyleId("DejaVuStyle");
                document.AddParagraph("Fallback paragraph");

                document.Save(openWord);
            }
        }
    }
}
