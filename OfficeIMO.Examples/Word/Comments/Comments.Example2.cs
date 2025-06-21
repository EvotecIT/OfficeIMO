using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Comments {
        internal static void Example_RemoveCommentsAndTrack(string folderPath, bool openWord) {
            Console.WriteLine("[*] Demonstrating comment lifecycle");
            string filePath = System.IO.Path.Combine(folderPath, "Comments Lifecycle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.TrackComments = true;
                var paragraph = document.AddParagraph("Paragraph with comment");
                paragraph.AddComment("John Doe", "JD", "My comment");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                if (document.Comments.Count > 0) {
                    document.Comments[0].Remove();
                }
                document.Save(openWord);
            }
        }
    }
}
