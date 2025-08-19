using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentTextBuilder(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent text");
            string filePath = Path.Combine(folderPath, "FluentTextBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Paragraph(p => p.Text("Hello")
                        .Text(" World", t => t.Bold().Italic().Color("ff0000")));
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
