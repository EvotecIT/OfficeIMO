using System;
using System.IO;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentImagesMixed(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with mixed images");
            string filePath = Path.Combine(folderPath, "FluentMixedImages.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Company report with ");
                    paragraph1.AddImage(Path.Combine(imagesPath, "EvotecLogo.png"), 24, 24, description: "");
                    paragraph1 = paragraph1.AddText(" logo");
                    var image3 = document.AddParagraph().AddImage(Path.Combine(imagesPath, "Kulek.jpg"));
                    image3.Width = 500;
                    image3.WrapText = WordImageTextWrapping.Square;
                    image3.Paragraph.SetAlignment(WordParagraphAlignment.Center);
                }

                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}