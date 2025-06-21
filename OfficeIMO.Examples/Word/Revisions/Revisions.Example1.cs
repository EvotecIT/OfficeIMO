using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Revisions {
        internal static void Example_TrackedChanges(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating tracked changes");
            string filePath = Path.Combine(folderPath, "TrackedChangesExample.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Original text:");
                paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Inserted text", "Codex");
                paragraph.AddDeletedText("Deleted text", "Codex");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AcceptRevisions();
                document.Save(openWord);
            }
        }
    }
}
