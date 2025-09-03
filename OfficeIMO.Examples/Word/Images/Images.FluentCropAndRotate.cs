using System;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageFluentCropAndRotate(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with cropped and rotated image using fluent API");
            string filePath = System.IO.Path.Combine(folderPath, "ImageFluentCropRotate.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            document.AsFluent()
                .Image(i => i.Add(System.IO.Path.Combine(imagePaths, "Kulek.jpg"))
                    .Size(200, 200)
                    .Crop(1, 1, 1, 1)
                    .Rotate(45))
                .End();
            document.Save(openWord);
        }
    }
}

