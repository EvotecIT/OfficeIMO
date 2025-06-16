using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddMultipleShapes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple shapes");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithMultipleShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("First shape");
                p1.AddShape(80, 40, "#00FF00");

                var p2 = document.AddParagraph("Second shape");
                p2.AddShape(50, 60, "#0000FF");

                document.Save(openWord);
            }
        }
    }
}
