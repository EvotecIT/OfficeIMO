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
                    .Info(i => i.SetTitle("Fluent Document")
                        .SetAuthor("Fluent Author")
                        .SetSubject("Fluent Subject")
                        .SetKeywords("fluent, api")
                        .SetComments("Created via fluent API")
                        .SetCustomProperty("Reviewed", true))
                    .Section(s => s.AddSection())
                    .Paragraph(p => p.Text("Hello from fluent API"));
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
