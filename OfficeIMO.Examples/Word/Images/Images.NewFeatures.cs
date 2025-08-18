using System;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

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
            paragraph1.Image.AlphaInversionColor = Color.Red;
            paragraph1.Image.BlackWhiteThreshold = 60;
            paragraph1.Image.BlurRadius = 5000;
            paragraph1.Image.BlurGrow = true;
            paragraph1.Image.ColorChangeFrom = Color.Parse("#97E4FE");
            paragraph1.Image.ColorChangeTo = Color.Parse("#FF3399");
            paragraph1.Image.ColorReplacement = Color.Lime;
            paragraph1.Image.DuotoneColor1 = Color.Black;
            paragraph1.Image.DuotoneColor2 = Color.White;
            paragraph1.Image.GrayScale = true;
            paragraph1.Image.LuminanceBrightness = 65;
            paragraph1.Image.LuminanceContrast = 30;
            paragraph1.Image.TintAmount = 50;
            paragraph1.Image.TintHue = 300;

            var paragraph2 = document.AddParagraph("Fit image");
            paragraph2.AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 50);
            paragraph2.Image.FillMode = ImageFillMode.Fit;

            var paragraph3 = document.AddParagraph("Centered image");
            paragraph3.AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 50);
            paragraph3.Image.FillMode = ImageFillMode.Center;

            var paragraph4 = document.AddParagraph("Linked image from web");
            paragraph4.AddImage(new Uri("http://example.com/logo.png"), 100, 100);

            document.Save(openWord);
        }
    }
}
