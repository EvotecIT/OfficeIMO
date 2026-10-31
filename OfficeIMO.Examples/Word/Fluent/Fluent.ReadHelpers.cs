using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentReadHelpers(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Using fluent read helpers");
            string filePath = Path.Combine(folderPath, "FluentReadHelpers.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("First");
                    var paragraph2 = document.AddParagraph();
                    paragraph2 = paragraph2.AddText("Second");
                    var paragraph3 = document.AddParagraph();
                    paragraph3 = paragraph3.AddText("Third");
                }

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath))
            {
                foreach (var paragraph in document.Find("Second"))
                    Console.WriteLine($"Found: {paragraph.Text}");
                foreach (var paragraph in document.Find(new System.Text.RegularExpressions.Regex("Sec.*")).BodyParagraphs)
                    Console.WriteLine($"Regex found: {paragraph.Text}");
                int totalRuns = 0;
                foreach (var paragraph in document.Paragraphs)
                    totalRuns += paragraph.GetRuns().Count();
                Console.WriteLine($"Total runs: {totalRuns}");
                int total = document.Paragraphs.Count(p => p.Text?.Contains("ir") == true);
                Console.WriteLine($"Paragraphs containing 'ir': {total}");
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}