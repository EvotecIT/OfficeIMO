using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageTransparencySimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Adding image with transparency");
            string filePath = Path.Combine(folderPath, "ImageTransparencySimple.docx");
            string imagePaths = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddImage(Path.Combine(imagePaths, "Kulek.jpg"), 100, 100);
                paragraph.Image.Transparency = 30;
                document.Save(openWord);
            }
        }

        internal static void Example_ImageTransparencyAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Modifying transparency in existing document");
            string templatesPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string filePath = Path.Combine(folderPath, "ImageTransparencyAdvanced.docx");
            File.Copy(Path.Combine(templatesPath, "BasicDocumentWithImages.docx"), filePath, true);

            using (WordDocument document = WordDocument.Load(filePath, false)) {
                document.Images[0].Transparency = 75;
                document.Save(openWord);
            }
        }
    }
}
