using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddBasicShape(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a basic rectangle shape");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithShape.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with red rectangle");
                paragraph.AddShape(100, 50, Color.Red);
                document.Save(openWord);
            }
        }
    }
}
