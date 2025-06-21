using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddEllipseAndPolygon(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with ellipse and polygon shapes");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithEllipsePolygon.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with shapes");
                WordShape.AddEllipse(paragraph, 80, 40, Color.Red);
                WordShape.AddPolygon(paragraph, "0,0 40,0 40,40 0,40", Color.Lime, Color.Blue);
                document.Save(openWord);
            }
        }
    }
}
