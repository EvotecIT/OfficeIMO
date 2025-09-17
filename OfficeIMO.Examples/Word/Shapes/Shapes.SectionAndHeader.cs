using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddShapesInSectionAndHeader(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shapes in section and header");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithSectionAndHeaderShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var section = document.Sections[0];
                section.AddShape(ShapeType.Rectangle, 50, 25, Color.Red, Color.Black);
                section.AddShape(ShapeType.RoundedRectangle, 40, 20, Color.Yellow, Color.Purple, 1, arcSize: 0.3);
                section.AddShapeDrawing(ShapeType.Ellipse, 40, 40);

                section.AddHeadersAndFooters();
                var sectionHeader = GetSectionHeaderOrThrow(section);
                sectionHeader.AddShape(ShapeType.Rectangle, 30, 20, Color.Blue, Color.Black);
                sectionHeader.AddShape(ShapeType.RoundedRectangle, 25, 15, Color.Green, Color.Black, 1, arcSize: 0.3);
                sectionHeader.AddShapeDrawing(ShapeType.Ellipse, 20, 20);

                document.Save(openWord);
            }
        }
    }
}
