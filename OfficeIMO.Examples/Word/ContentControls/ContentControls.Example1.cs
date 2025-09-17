using System;
using System.IO;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class ContentControls {
        internal static void Example_AddContentControl(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a content control");
            string filePath = Path.Combine(folderPath, "DocumentWithContentControl.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var control = document.AddStructuredDocumentTag("Sample text", "ExampleAlias", "ExampleTag");

                Console.WriteLine($"Alias: {control.Alias}");
                Console.WriteLine($"Text: {control.Text}");

                control.Text = "Updated text";
                document.Save(openWord);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var loaded = Guard.NotNull(document.GetStructuredDocumentTagByTag("ExampleTag"), "Structured document tag 'ExampleTag' was not found.");
                Console.WriteLine($"Loaded text: {loaded.Text}");
                document.Save(openWord);
            }
        }
    }
}
