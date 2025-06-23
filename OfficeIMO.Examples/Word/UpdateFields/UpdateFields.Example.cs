using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class UpdateFieldsSample {
        internal static void Example_UpdateFields(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating UpdateFields vs UpdateFieldsOnOpen");
            string filePath = System.IO.Path.Combine(folderPath, "Document with Updated Fields.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                // Option 1: update fields when the document is opened in Word
                document.Settings.UpdateFieldsOnOpen = true;

                document.AddParagraph("Page 1").AddPageNumber(includeTotalPages: true);
                document.AddPageBreak();
                document.AddParagraph("Page 2");
                document.AddTableOfContent();

                // Option 2: update fields immediately in code
                document.UpdateFields();
                document.Save(openWord);
            }
        }
    }
}
