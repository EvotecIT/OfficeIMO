using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageCroppingBasic(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with cropped image");
            var filePath = System.IO.Path.Combine(folderPath, "ImageCroppingBasic.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph("Cropped picture below:");
            paragraph.AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 200, 200);

            paragraph.Image.CropTop = 1000;
            paragraph.Image.CropBottom = 1000;
            paragraph.Image.CropLeft = 1000;
            paragraph.Image.CropRight = 1000;

            document.Save(openWord);
        }
    }
}
