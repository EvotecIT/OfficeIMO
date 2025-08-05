using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static partial class Fonts {
        public static void Example_FontResolverFallback(string folderPath) {
            Console.WriteLine("[*] Demonstrating font resolver fallback");
            string filePath = Path.Combine(folderPath, "DocumentFontFallback.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                string resolved = FontResolver.Resolve("MissingFont")!;
                document.AddParagraph($"Paragraph using {resolved} font").SetFontFamily(resolved);
                document.Save(false);
            }
        }
    }
}
