using System;
using System.IO;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static async Task Example_FluentImagesFromUrlAsync(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with image from URL asynchronously");
            string filePath = Path.Combine(folderPath, "FluentImageFromUrlAsync.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                var image = await document.AddImageFromUrlAsync("https://raw.githubusercontent.com/EvotecIT/OfficeIMO/master/OfficeIMO.Examples/Images/Kulek.jpg");
                image.WrapText = WordImageTextWrapping.Tight;
                image.Paragraph.ParagraphAlignment = WordParagraphAlignment.Right;
                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}