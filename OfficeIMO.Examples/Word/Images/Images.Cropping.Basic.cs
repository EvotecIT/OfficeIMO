using System;
using OfficeIMO.Examples.Utils;
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
            var image = Guard.NotNull(paragraph.Image, "Paragraph should contain an image for cropping.");

            image.CropTopCentimeters = 1;
            image.CropBottomCentimeters = 1;
            image.CropLeftCentimeters = 1;
            image.CropRightCentimeters = 1;

            document.Save(openWord);
        }
    }
}
