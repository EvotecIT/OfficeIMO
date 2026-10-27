using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentImagesMultiple(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with multiple block images");
            string filePath = Path.Combine(folderPath, "FluentMultipleImages.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var image2 = document.AddParagraph().AddImage(Path.Combine(imagesPath, "PrzemyslawKlysAndKulkozaurr.jpg"));
                    image2.Width = 200;
                    image2.Paragraph.SetAlignment(WordParagraphAlignment.Left);
                    var image3 = document.AddParagraph().AddImage(Path.Combine(imagesPath, "Kulek.jpg"));
                    image3.SetMaximumWidth(OfficeLength.Pixels(300));
                    image3.Paragraph.SetAlignment(WordParagraphAlignment.Right);
                }

                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}