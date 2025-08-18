using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageCompressionQuality(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating image compression quality");
            string filePath = System.IO.Path.Combine(folderPath, "ImageCompressionQuality.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph("Image with compression quality");
            paragraph.AddImage(System.IO.Path.Combine(imagePaths, "Kulek.jpg"), 100, 100, WrapTextImage.BehindText);
            paragraph.Image.CompressionQuality = BlipCompressionValues.HighQualityPrint;
            document.Save(openWord);
        }
    }
}
