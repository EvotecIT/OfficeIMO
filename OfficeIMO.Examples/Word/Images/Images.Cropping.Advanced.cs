using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class Images
    {
        internal static void Example_ImageCroppingAdvanced(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with advanced cropped image");
            var filePath = System.IO.Path.Combine(folderPath, "ImageCroppingAdvanced.docx");
            var imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph("Advanced crop with shape:");
            var image = paragraph.AddImage(System.IO.Path.Combine(imagePaths, "PrzemyslawKlysAndKulkozaurr.jpg"), OfficeIMO.OfficeLength.Pixels(300), OfficeIMO.OfficeLength.Pixels(300), options: new OfficeIMO.Word.WordImageOptions { Wrapping = WordImageTextWrapping.Square });
            image.Shape = OfficePresetShapeType.Cube;
            image.CropTop = OfficeIMO.OfficeLength.Centimeters(2);
            image.CropBottom = OfficeIMO.OfficeLength.Centimeters(1.5);
            image.CropLeft = OfficeIMO.OfficeLength.Centimeters(0.5);
            image.CropRight = OfficeIMO.OfficeLength.Centimeters(0.5);
            document.Save();
            if (openWord)
                document.OpenInApplication();
        }
    }
}