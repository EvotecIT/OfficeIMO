using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentImagesBlockLevel(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a block-level image");
            string filePath = Path.Combine(folderPath, "FluentBlockImage.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(img => img
                        .Add(Path.Combine(imagesPath, "EvotecLogo.png"))
                        .Size(120, 120)
                        .Wrap(WrapTextImage.Square)
                        .Align(HorizontalAlignment.Center))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
