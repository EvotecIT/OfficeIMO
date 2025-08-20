using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentImagesMultiple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple block images");
            string filePath = Path.Combine(folderPath, "FluentMultipleImages.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(img => img
                        .Add(Path.Combine(imagesPath, "PrzemyslawKlysAndKulkozaurr.jpg"))
                            .Size(200)
                            .Align(HorizontalAlignment.Left)
                        .Add(Path.Combine(imagesPath, "Kulek.jpg"))
                            .MaxWidth(300)
                            .Align(HorizontalAlignment.Right))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
