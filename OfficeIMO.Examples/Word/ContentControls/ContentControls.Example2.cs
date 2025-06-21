using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_MultipleContentControls(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple content controls");
            string filePath = Path.Combine(folderPath, "DocumentWithMultipleContentControls.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddStructuredDocumentTag("First", "Alias1", "Tag1");
                document.AddStructuredDocumentTag("Second", "Alias2", "Tag2");
                document.AddStructuredDocumentTag("Third", "Alias3", "Tag3");
                Console.WriteLine("Controls: " + document.StructuredDocumentTags.Count);
                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                foreach (var control in document.StructuredDocumentTags) {
                    Console.WriteLine(control.Tag + ": " + control.Text);
                }
                document.Save(openWord);
            }
        }
    }
}
