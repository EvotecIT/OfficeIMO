using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentImagesInline(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with inline image");
            string filePath = Path.Combine(folderPath, "FluentInlineImage.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p
                        .Text("Here is an inline icon ")
                        .InlineImage(Path.Combine(imagesPath, "EvotecLogo.png"), widthPx: 16, heightPx: 16, alt: "icon")
                        .Text(" followed by text."))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
