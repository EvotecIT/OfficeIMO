using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fonts {
        public static void Example_EmbeddedAndBuiltinFonts(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document mixing builtin and embedded fonts");
            string filePath = Path.Combine(folderPath, "DocumentMixedFonts.docx");
            string fontPath = Path.Combine(templatesPath, "DejaVuSans.ttf");
            if (!File.Exists(fontPath)) {
                Console.WriteLine($"Font file not found: {fontPath}");
                return;
            }

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.EmbedFont(fontPath);

                document.AddParagraph("Paragraph in builtin Arial font.")
                    .SetFontFamily("Arial");

                document.AddParagraph("Paragraph in embedded DejaVu Sans font.")
                    .SetFontFamily("DejaVu Sans");

                var paragraph = document.AddParagraph("Mix of builtin and embedded fonts within one paragraph: ");
                paragraph.AddText("Arial text, ").SetFontFamily("Arial");
                paragraph.AddText("DejaVu text").SetFontFamily("DejaVu Sans");

                document.Save(openWord);
            }
        }
    }
}