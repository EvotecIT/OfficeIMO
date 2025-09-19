using System;
using OfficeIMO.Examples.Utils;
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
            var paragraph1Image = paragraph1.InsertImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 100);
            paragraph1Image.FillMode = ImageFillMode.Tile;
            paragraph1Image.UseLocalDpi = true;
            paragraph1Image.Title = "Sample image";
            paragraph1Image.Hidden = false;
            paragraph1Image.PreferRelativeResize = true;
            paragraph1Image.NoChangeAspect = true;
            paragraph1Image.FixedOpacity = 80;
            paragraph1Image.AlphaInversionColor = Color.Red;
            paragraph1Image.BlackWhiteThreshold = 60;
            paragraph1Image.BlurRadius = 5000;
            paragraph1Image.BlurGrow = true;
            paragraph1Image.ColorChangeFrom = Color.Parse("#97E4FE");
            paragraph1Image.ColorChangeTo = Color.Parse("#FF3399");
            paragraph1Image.ColorReplacement = Color.Lime;
            paragraph1Image.DuotoneColor1 = Color.Black;
            paragraph1Image.DuotoneColor2 = Color.White;
            paragraph1Image.GrayScale = true;
            paragraph1Image.LuminanceBrightness = 65;
            paragraph1Image.LuminanceContrast = 30;
            paragraph1Image.TintAmount = 50;
            paragraph1Image.TintHue = 300;

            var paragraph2 = document.AddParagraph("Fit image");
            paragraph2.InsertImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 50).FillMode = ImageFillMode.Fit;

            var paragraph3 = document.AddParagraph("Centered image");
            paragraph3.InsertImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 50).FillMode = ImageFillMode.Center;

            var paragraph4 = document.AddParagraph("Linked image from web");
            paragraph4.InsertImage(new Uri("http://example.com/logo.png"), 100, 100);

            document.Save(openWord);
        }
    }
}
