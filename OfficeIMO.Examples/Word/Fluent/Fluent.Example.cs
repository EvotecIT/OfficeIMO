using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentDocument(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent API");
            string filePath = Path.Combine(folderPath, "FluentDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Info(i => i.Title("Fluent Document")
                        .Author("Fluent Author")
                        .Subject("Fluent Subject")
                        .Keywords("fluent, api")
                        .Comments("Created via fluent API")
                        .Custom("Reviewed", true))
                    .Section(s => s.New())
                    .Paragraph(p => p.Text("Hello from fluent API"))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
