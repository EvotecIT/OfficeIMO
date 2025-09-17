using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageTransparencySimple(string folderPath, bool openWord) {
            Console.WriteLine("[*] Adding image with transparency");
            string filePath = Path.Combine(folderPath, "ImageTransparencySimple.docx");
            string imagePaths = Path.Combine(Directory.GetCurrentDirectory(), "Images");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                // AddImage returns the paragraph (fluent API). Set transparency on the last added image.
                paragraph.AddImage(Path.Combine(imagePaths, "Kulek.jpg"), 100, 100);
                var images = document.Images;
                var insertedImage = Guard.GetRequiredItem(images, images.Count - 1, "Document should contain the inserted image before setting transparency.");
                insertedImage.Transparency = 30;
                document.Save(openWord);
            }
        }

        internal static void Example_ImageTransparencyAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Modifying transparency in existing document");
            string templatesPath = Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            string filePath = Path.Combine(folderPath, "ImageTransparencyAdvanced.docx");
            File.Copy(Path.Combine(templatesPath, "BasicDocumentWithImages.docx"), filePath, true);

            using (WordDocument document = WordDocument.Load(filePath, false)) {
                var loadedImages = document.Images;
                var firstImage = Guard.GetRequiredItem(loadedImages, 0, "Template document should contain at least one image before adjusting transparency.");
                firstImage.Transparency = 75;
                document.Save(openWord);
            }
        }
    }
}
