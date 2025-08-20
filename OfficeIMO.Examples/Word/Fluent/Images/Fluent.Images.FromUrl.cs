using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentImagesFromUrl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with image from URL");
            string filePath = Path.Combine(folderPath, "FluentImageFromUrl.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Image(img => img
                        .AddFromUrl("https://raw.githubusercontent.com/EvotecIT/OfficeIMO/master/OfficeIMO.Examples/Images/Kulek.jpg")
                        .Size(400)
                        .Align(HorizontalAlignment.Center))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
