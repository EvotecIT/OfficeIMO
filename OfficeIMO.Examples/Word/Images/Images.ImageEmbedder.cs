using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_ImageEmbedderHelper(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with ImageEmbedder helper");

            string filePath = Path.Combine(folderPath, "ImageEmbedder.docx");

            using WordDocument document = WordDocument.Create(filePath);
            string imagePath = Path.Combine("Assets", "OfficeIMO.png");
            document.AddParagraph().AddImage(imagePath);
            document.Save();

            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
