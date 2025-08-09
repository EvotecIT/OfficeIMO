using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageSaveToFileExclusive(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating exclusive access when saving images");
            string filePath = Path.Combine(folderPath, "ImageSaveExclusive.docx");
            string imagePaths = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using var document = WordDocument.Create(filePath);
            var paragraph = document.AddParagraph();
            paragraph.AddImage(Path.Combine(imagePaths, "Kulek.jpg"), 50, 50);

            string fileToSave = Path.Combine(folderPath, "LockedImage.jpg");
            using (var lockStream = new FileStream(fileToSave, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite)) {
                try {
                    paragraph.Image.SaveToFile(fileToSave);
                } catch (IOException) {
                    Console.WriteLine("[!] Unable to save image while file is locked");
                }
            }

            paragraph.Image.SaveToFile(fileToSave);
            document.Save(openWord);
        }
    }
}
