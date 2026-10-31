using System;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddBasicShape(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a basic rectangle shape");
            string filePath = System.IO.Path.Combine(folderPath, "BasicDocumentWithShape.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with red rectangle");
                var shp = paragraph.AddShape(OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(50), Color.Red);
                // Demonstrate property setters
                shp.Stroked = true;
                shp.StrokeColor = Color.Black;
                shp.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                // Add a couple more simple shapes, spaced vertically
                var p2 = document.AddParagraph("Ellipse below");
                p2.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(50), Color.Orange, Color.Black, OfficeIMO.OfficeLength.Points(1.25));

                var p3 = document.AddParagraph("Line below");
                p3.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(0), Color.Blue, OfficeIMO.OfficeLength.Points(2));
                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }
    }
}
