using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentReadHelpers(string folderPath, bool openWord) {
            Console.WriteLine("[*] Using fluent read helpers");
            string filePath = Path.Combine(folderPath, "FluentReadHelpers.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraphs.AddParagraph("First")
                    .Paragraphs.AddParagraph("Second")
                    .Paragraphs.AddParagraph("Third");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AsFluent().Find("Second", p => Console.WriteLine($"Found: {p.Paragraph?.Text}"));
                int total = document.AsFluent().Select(p => p.Paragraph?.Text.Contains("ir") == true).Count();
                Console.WriteLine($"Paragraphs containing 'ir': {total}");
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
