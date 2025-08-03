using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Revisions {
        internal static void Example_ConvertRevisionsToMarkup(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating converting revisions to markup");
            string filePath = Path.Combine(folderPath, "TrackedChangesMarkup.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                paragraph.AddInsertedText("Inserted text", "Codex");
                paragraph.AddDeletedText("Deleted text", "Codex");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.ConvertRevisionsToMarkup();
                document.Save(openWord);
            }
        }
    }
}

