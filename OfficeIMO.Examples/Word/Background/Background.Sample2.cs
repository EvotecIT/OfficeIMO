using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Background {
        internal static void Example_BackgroundImageAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with background image from stream");
            string filePath = Path.Combine(folderPath, "BackgroundImageAdvanced.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            string imagePath = Path.Combine(imagesPath, "BackgroundImage.png");

            using var stream = File.OpenRead(imagePath);
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Background.SetImage(stream, "BackgroundImage.png", 600, 800);
                document.AddParagraph("Advanced content");
                document.Save(openWord);
            }
        }
    }
}
