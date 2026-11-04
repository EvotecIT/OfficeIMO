using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentParagraphContent(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with links, tabs, and breaks using fluent API");
            string filePath = Path.Combine(folderPath, "FluentParagraphContent.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Before");
                    paragraph1.AddTab();
                    paragraph1 = paragraph1.AddText("After");
                    var paragraph2 = document.AddParagraph();
                    paragraph2.AddHyperLink("Example", new Uri("https://example.com"), false);
                    var paragraph3 = document.AddParagraph();
                    paragraph3 = paragraph3.AddText("Line1");
                    paragraph3.AddBreak();
                    paragraph3 = paragraph3.AddText("Line2");
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}