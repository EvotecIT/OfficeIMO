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

            paragraph.Image.CropTopCentimeters = 1;
            paragraph.Image.CropBottomCentimeters = 1;
            paragraph.Image.CropLeftCentimeters = 1;
            paragraph.Image.CropRightCentimeters = 1;

            document.Save(openWord);
        }
    }
}
