using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageCroppingAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced cropped image");
            var filePath = System.IO.Path.Combine(folderPath, "ImageCroppingAdvanced.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph("Advanced crop with shape:");
            var image = paragraph.InsertImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 300, 300, WrapTextImage.Square);

            image.Shape = ShapeTypeValues.Cube;
            image.CropTopCentimeters = 2;
            image.CropBottomCentimeters = 1.5;
            image.CropLeftCentimeters = 0.5;
            image.CropRightCentimeters = 0.5;

            document.Save(openWord);
        }
    }
}
