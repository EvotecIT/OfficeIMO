using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fonts {
        public static void Example_EmbedFont(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with embedded font");
            string filePath = Path.Combine(folderPath, "DocumentWithEmbeddedFont.docx");
            string fontPath = Path.Combine(templatesPath, "DejaVuSans.ttf");
            if (!File.Exists(fontPath)) {
                Console.WriteLine($"Font file not found: {fontPath}");
                return;
            }

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("This document uses an embedded font.");
                document.EmbedFont(fontPath);
                document.Save(openWord);
            }
        }
    }
}
