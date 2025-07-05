using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Background {
        internal static void Example_BackgroundImageSimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with background image");
            string filePath = Path.Combine(folderPath, "BackgroundImageSimple.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            string imagePath = Path.Combine(imagesPath, "BackgroundImage.png");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Background.SetImage(imagePath);
                document.AddParagraph("Content on top of image");
                document.Save(openWord);
            }
        }
    }
}
