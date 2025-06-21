using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_AddContentControl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a content control");
            string filePath = Path.Combine(folderPath, "DocumentWithContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var control = document.AddStructuredDocumentTag("Sample text", "ExampleAlias");

                Console.WriteLine($"Alias: {control.Alias}");
                Console.WriteLine($"Text: {control.Text}");

                control.Text = "Updated text";
                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Loaded text: {document.StructuredDocumentTags[0].Text}");
                document.Save(openWord);
            }
        }
    }
}
