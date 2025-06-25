using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_PictureBulletList(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with picture bullets");
            string filePath = Path.Combine(folderPath, "Document picture bullets.docx");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "Kulek.jpg");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var list = document.AddPictureBulletList(imagePath);
                list.AddItem("Picture item 1");
                list.AddItem("Picture item 2");
                document.Save(openWord);
            }
        }
    }
}
