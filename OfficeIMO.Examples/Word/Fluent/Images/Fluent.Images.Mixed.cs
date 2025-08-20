using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentImagesMixed(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with mixed images");
            string filePath = Path.Combine(folderPath, "FluentMixedImages.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p
                        .Text("Company report with ")
                        .InlineImage(Path.Combine(imagesPath, "EvotecLogo.png"), widthPx: 24, heightPx: 24)
                        .Text(" logo"))
                    .Image(img => img
                        .Add(Path.Combine(imagesPath, "Kulek.jpg"))
                        .Size(500)
                        .Wrap(WrapTextImage.Square)
                        .Align(HorizontalAlignment.Center))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
