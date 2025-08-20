using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentParagraphContent(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with links, tabs, and breaks using fluent API");
            string filePath = Path.Combine(folderPath, "FluentParagraphContent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Before").Tab().Text("After"))
                    .Paragraph(p => p.Link("https://example.com", "Example"))
                    .Paragraph(p => p.Text("Line1").Break().Text("Line2"))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
