using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageNewFeatures(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating image fill modes and external images");
            string filePath = System.IO.Path.Combine(folderPath, "ImageNewFeatures.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph1 = document.AddParagraph("Tiled image with local DPI");
            paragraph1.AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 100);
            paragraph1.Image.FillMode = ImageFillMode.Tile;
            paragraph1.Image.UseLocalDpi = true;
            paragraph1.Image.Title = "Sample image";
            paragraph1.Image.Hidden = false;
            paragraph1.Image.PreferRelativeResize = true;
            paragraph1.Image.NoChangeAspect = true;
            paragraph1.Image.FixedOpacity = 80;

            var paragraph2 = document.AddParagraph("Linked image from web");
            paragraph2.AddImage(new Uri("http://example.com/logo.png"), 100, 100);

            document.Save(openWord);
        }
    }
}
