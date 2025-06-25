using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_PictureBulletListAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating advanced document with picture bullets");
            string filePath = Path.Combine(folderPath, "Document picture bullets advanced.docx");
            string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Images", "Kulek.jpg");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Shopping list:");

                var list = document.AddPictureBulletList(imagePath);
                list.AddItem("Milk");
                list.AddItem("Bread");
                list.AddItem("Butter");

                document.AddParagraph("End of list.");
                document.Save(openWord);
            }
        }
    }
}
