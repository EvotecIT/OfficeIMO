using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentHeadersAndFooters(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent headers and footers");
            string filePath = Path.Combine(folderPath, "FluentHeadersAndFooters.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    document.DifferentFirstPage = true;
                    document.DifferentOddAndEvenPages = true;
                    document.HeaderDefaultOrCreate.AddParagraph("Default header");
                    var paragraph1 = document.HeaderDefaultOrCreate.AddParagraph();
                    paragraph1 = paragraph1.AddText("Second header paragraph");
                    document.HeaderDefaultOrCreate.AddParagraph().AddImage(Path.Combine(imagesPath, "Kulek.jpg"), OfficeIMO.OfficeLength.Pixels(50), OfficeIMO.OfficeLength.Pixels(50));
                    document.HeaderDefaultOrCreate.AddTable(2, 2);
                    document.HeaderFirstOrCreate.AddParagraph("First page header");
                    document.HeaderEvenOrCreate.AddParagraph("Even page header");
                    document.FooterDefaultOrCreate.AddParagraph("Default footer");
                    document.FooterFirstOrCreate.AddParagraph("First page footer");
                    document.FooterEvenOrCreate.AddParagraph("Even page footer");
                    var paragraph2 = document.AddParagraph();
                    paragraph2 = paragraph2.AddText("Body paragraph");
                    var section3 = document.Sections.Last();
                    section3 = document.AddSection();
                    var paragraph4 = document.AddParagraph();
                    paragraph4 = paragraph4.AddText("Second section paragraph");
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}