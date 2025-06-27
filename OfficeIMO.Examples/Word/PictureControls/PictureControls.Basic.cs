using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PictureControls {
        internal static void Example_BasicPictureControl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a picture content control");
            string filePath = Path.Combine(folderPath, "DocumentWithPictureControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var imagePath = Path.Combine(folderPath, "Images", "Kulek.jpg");
                document.AddParagraph().AddPictureControl(imagePath, 100, 100, "Pic", "PicTag");
                document.Save(openWord);
            }
        }
    }
}
