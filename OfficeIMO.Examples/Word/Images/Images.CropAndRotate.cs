using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class Images
    {
        internal static void Example_ImageCropAndRotate(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with a cropped and rotated image");
            string filePath = System.IO.Path.Combine(folderPath, "ImageCropRotate.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            using var document = WordDocument.Create(filePath);
            {
                var image2 = document.AddParagraph().AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"));
                image2.Width = OfficeIMO.OfficeLength.Pixels(200);
                image2.Height = OfficeIMO.OfficeLength.Pixels(200);
                image2.SetCrop(OfficeIMO.OfficeLength.Centimeters(1), OfficeIMO.OfficeLength.Centimeters(1), OfficeIMO.OfficeLength.Centimeters(1), OfficeIMO.OfficeLength.Centimeters(1));
                image2.Rotation = 45;
            }

            document.Save();
            if (openWord)
                document.OpenInApplication();
        }
    }
}