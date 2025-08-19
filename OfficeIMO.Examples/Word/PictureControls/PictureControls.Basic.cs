using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PictureControls {
        internal static void Example_BasicPictureControl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a picture content control");
            string filePath = Path.Combine(folderPath, "DocumentWithPictureControl.docx");
            string imagePaths = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var imagePath = Path.Combine(imagePaths, "Kulek.jpg");
                document.AddParagraph().AddPictureControl(imagePath, 100, 100, "Pic", "PicTag");
                document.Save(openWord);
            }
        }
    }
}
