using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentImagesInline(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with inline image");
            string filePath = Path.Combine(folderPath, "FluentInlineImage.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Here is an inline icon ");
                    paragraph1.AddImage(Path.Combine(imagesPath, "EvotecLogo.png"), OfficeIMO.OfficeLength.Pixels(16), OfficeIMO.OfficeLength.Pixels(16), options: new OfficeIMO.Word.WordImageOptions { Description = "icon" });
                    paragraph1 = paragraph1.AddText(" followed by text.");
                }

                document.Save();
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}