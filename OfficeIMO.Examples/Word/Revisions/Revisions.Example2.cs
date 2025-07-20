using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Revisions {
        internal static void Example_TrackChangesToggle(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating TrackChanges toggle");
            string filePath = Path.Combine(folderPath, "TrackChangesToggle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.TrackChanges = true;
                document.AddParagraph("Tracking enabled");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.TrackChanges = false;
                document.AddParagraph("Tracking disabled");
                document.Save(openWord);
            }
        }
    }
}
