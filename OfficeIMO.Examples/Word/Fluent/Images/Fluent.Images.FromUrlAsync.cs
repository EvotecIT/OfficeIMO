using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static async Task Example_FluentImagesFromUrlAsync(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with image from URL asynchronously");
            string filePath = Path.Combine(folderPath, "FluentImageFromUrlAsync.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                await document.AsFluent()
                    .ImageAsync(async img => {
                        await img.AddFromUrlAsync("https://raw.githubusercontent.com/EvotecIT/OfficeIMO/master/OfficeIMO.Examples/Images/Kulek.jpg");
                        img.Wrap(WrapTextImage.Tight).Align(HorizontalAlignment.Right);
                    });
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
