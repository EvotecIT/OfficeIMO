using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Comments {
        internal static void Example_PlayingWithComments(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with comments");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with comments.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test Section");

                document.Paragraphs[0].AddComment("Przemysław", "PK", "This is my comment");


                document.AddParagraph("Test Section - another line");

                document.Paragraphs[1].AddComment("Przemysław", "PK", "More comments");

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.Save(true);
            }
        }

    }
}
