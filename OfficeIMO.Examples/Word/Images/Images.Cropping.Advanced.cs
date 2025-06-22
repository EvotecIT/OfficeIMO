using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageCroppingAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced cropped image");
            var filePath = System.IO.Path.Combine(folderPath, "ImageCroppingAdvanced.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph("Advanced crop with shape:");
            paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), 300, 300, WrapTextImage.Square);

            paragraph.Image.Shape = ShapeTypeValues.Cube;
            paragraph.Image.CropTop = 2000;
            paragraph.Image.CropBottom = 1500;
            paragraph.Image.CropLeft = 500;
            paragraph.Image.CropRight = 500;

            document.Save(openWord);
        }
    }
}
