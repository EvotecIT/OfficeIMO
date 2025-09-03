using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DocumentProperties {
        public static void Example_FluentDocumentProperties(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent document properties");

            string filePath = Path.Combine(folderPath, "Fluent Document Properties.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Info(i => i.Title("Title")
                        .Author("Author")
                        .Subject("Subject")
                        .Keywords("k1, k2")
                        .Comments("Some comments")
                        .Category("Category")
                        .Company("Evotec")
                        .Manager("Manager1")
                        .LastModifiedBy("John")
                        .Revision("1.0"))
                    .Paragraph(p => p.Text("Test"));

                document.Save(openWord);
            }
        }
    }
}

