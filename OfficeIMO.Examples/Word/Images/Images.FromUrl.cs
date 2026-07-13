using System;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static async Task Example_AddImageFromUrlAsync(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with image downloaded from URL");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithImageFromUrl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                await document.AddImageFromUrlAsync("https://via.placeholder.com/150", 150, 150);
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
