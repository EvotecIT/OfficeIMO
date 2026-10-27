using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentImagesBlockLevel(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with a block-level image");
            string filePath = Path.Combine(folderPath, "FluentBlockImage.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var image2 = document.AddParagraph().AddImage(Path.Combine(imagesPath, "EvotecLogo.png"));
                    image2.Width = 120;
                    image2.Height = 120;
                    image2.WrapText = WordImageTextWrapping.Square;
                    image2.Paragraph.SetAlignment(WordParagraphAlignment.Center);
                }

                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}