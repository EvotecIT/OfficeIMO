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
                var shp = paragraph.AddShape(100, 50, Color.Red);
                // Demonstrate property setters
                shp.Stroked = true;
                shp.StrokeColor = Color.Black;
                shp.StrokeWeight = 1.5;

                // Add a couple more simple shapes, spaced vertically
                var p2 = document.AddParagraph("Ellipse below");
                p2.AddShape(ShapeType.Ellipse, 80, 50, Color.Orange, Color.Black, 1.25);

                var p3 = document.AddParagraph("Line below");
                p3.AddLine(0, 0, 120, 0, Color.Blue, 2);
                document.Save(openWord);
            }
        }
    }
}
