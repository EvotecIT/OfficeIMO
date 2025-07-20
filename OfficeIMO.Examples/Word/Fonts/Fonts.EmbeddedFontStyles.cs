using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fonts {
        public static void Example_EmbeddedFontStyle(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document using embedded font in a custom style");
            string filePath = Path.Combine(folderPath, "DocumentEmbeddedFontStyle.docx");
            string fontPath = Path.Combine(templatesPath, "DejaVuSans.ttf");
            if (!File.Exists(fontPath)) {
                Console.WriteLine($"Font file not found: {fontPath}");
                return;
            }

            var style = new Style { Type = StyleValues.Paragraph, StyleId = "EmbeddedStyle" };
            style.Append(new StyleName { Val = "EmbeddedStyle" });
            var runProps = new StyleRunProperties();
            runProps.Append(new RunFonts { Ascii = "DejaVu Sans" });
            style.Append(runProps);
            WordParagraphStyle.RegisterCustomStyle("EmbeddedStyle", style);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.EmbedFont(fontPath);

                document.AddParagraph("Paragraph using custom embedded style").SetStyleId("EmbeddedStyle");
                document.AddParagraph("Paragraph using builtin Times New Roman").SetFontFamily("Times New Roman");

                document.Save(openWord);
            }
        }
    }
}