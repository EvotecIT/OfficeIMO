using System;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_CloneImage(string folderPath, bool openWord) {
            Console.WriteLine("[*] Cloning image in a document");
            var filePath = System.IO.Path.Combine(folderPath, "DocumentWithClonedImage.docx");
            var imagePath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images", "Kulek.jpg");

            using var document = WordDocument.Create(filePath);
            var paragraph1 = document.AddParagraph();
            paragraph1.AddImage(imagePath, 100, 100);

            var paragraph2 = document.AddParagraph();
            var image = Guard.NotNull(paragraph1.Image, "Source paragraph is expected to contain an image to clone.");
            image.Clone(paragraph2);

            document.Save(openWord);
        }
    }
}
