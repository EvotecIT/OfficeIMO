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
                    document.PageSetup.DifferentFirstPage = true;
                    document.PageSetup.DifferentOddAndEvenPages = true;
                    document.PageSetup.Header.AddParagraph("Default header");
                    var paragraph1 = document.PageSetup.Header.AddParagraph();
                    paragraph1 = paragraph1.AddText("Second header paragraph");
                    document.PageSetup.Header.AddParagraph().AddImage(Path.Combine(imagesPath, "Kulek.jpg"), OfficeIMO.OfficeLength.Pixels(50), OfficeIMO.OfficeLength.Pixels(50));
                    document.PageSetup.Header.AddTable(2, 2);
                    document.PageSetup.FirstPageHeader.AddParagraph("First page header");
                    document.PageSetup.EvenPageHeader.AddParagraph("Even page header");
                    document.PageSetup.Footer.AddParagraph("Default footer");
                    document.PageSetup.FirstPageFooter.AddParagraph("First page footer");
                    document.PageSetup.EvenPageFooter.AddParagraph("Even page footer");
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