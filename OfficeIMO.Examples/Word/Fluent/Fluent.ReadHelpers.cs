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
                    .Paragraph(p => p.Text("First"))
                    .Paragraph(p => p.Text("Second"))
                    .Paragraph(p => p.Text("Third"));
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AsFluent().Find("Second", p => Console.WriteLine($"Found: {p.Paragraph?.Text}"));
                document.AsFluent().FindRegex("Sec.*", p => Console.WriteLine($"Regex found: {p.Paragraph?.Text}"));
                int totalRuns = 0;
                document.AsFluent().ForEachRun(r => totalRuns++);
                Console.WriteLine($"Total runs: {totalRuns}");
                int total = document.AsFluent().Select(p => p.Paragraph?.Text.Contains("ir") == true).Count();
                Console.WriteLine($"Paragraphs containing 'ir': {total}");
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
