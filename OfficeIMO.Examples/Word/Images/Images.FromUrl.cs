using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_AddImageFromUrl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with image downloaded from URL");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithImageFromUrl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddImageFromUrl("https://via.placeholder.com/150", 150, 150);
                document.Save(openWord);
            }
        }
    }
}
