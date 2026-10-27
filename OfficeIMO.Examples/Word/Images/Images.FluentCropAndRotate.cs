using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class Images
    {
        internal static void Example_ImageFluentCropAndRotate(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with cropped and rotated image using fluent API");
            string filePath = System.IO.Path.Combine(folderPath, "ImageFluentCropRotate.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            using var document = WordDocument.Create(filePath);
            {
                var image2 = document.AddParagraph().AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"));
                image2.Width = 200;
                image2.Height = 200;
                image2.SetCropCentimeters(1, 1, 1, 1);
                image2.Rotation = (int)Math.Round((double)(45));
            }

            document.Save();
            if (openWord)
                document.OpenInApplication();
        }
    }
}