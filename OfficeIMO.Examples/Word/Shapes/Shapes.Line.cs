using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddLine(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a line shape");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithLine.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with line");
                paragraph.AddLine(0, 0, 100, 0, SixLabors.ImageSharp.Color.Red, 2);
                document.Save(openWord);
            }
        }
    }
}
