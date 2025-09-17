using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_AddMultipleShapes(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple shapes");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithMultipleShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // VML Rectangle with named color
                var p1 = document.AddParagraph("Rectangle (Green)");
                var rect = p1.AddShape(80, 40, SixLabors.ImageSharp.Color.Green);

                // VML Ellipse with stroke
                var p2 = document.AddParagraph("Ellipse (Orange fill, Black stroke)");
                var ellipse = p2.AddShape(ShapeType.Ellipse, 90, 50, SixLabors.ImageSharp.Color.Orange, SixLabors.ImageSharp.Color.Black, strokeWeightPt: 2);

                // VML Rounded Rectangle with arc size
                var p3 = document.AddParagraph("Rounded Rectangle (Blue)");
                var rrect = p3.AddShape(ShapeType.RoundedRectangle, 100, 50, SixLabors.ImageSharp.Color.Blue, SixLabors.ImageSharp.Color.Black, strokeWeightPt: 1.5, arcSize: 0.4);

                // VML Line with stroke color/weight
                var p4 = document.AddParagraph("Line (Red, 3pt)");
                var line = p4.AddLine(0, 0, 120, 0, SixLabors.ImageSharp.Color.Red, 3);

                // VML Polygon (triangle)
                var p5 = document.AddParagraph("Polygon (Triangle, Purple)");
                var poly = WordShape.AddPolygon(p5, "0,0 60,0 30,50 0,0", SixLabors.ImageSharp.Color.Purple, SixLabors.ImageSharp.Color.Black);

                // Absolute positioning for VML shapes
                rect.Left = 20; rect.Top = 20; rect.Rotation = 5;
                ellipse.Left = 160; ellipse.Top = 20;
                rrect.Left = 300; rrect.Top = 20; rrect.Rotation = -10;
                poly.Left = 20; poly.Top = 120;

                // A second row of VML shapes for visibility (own paragraph)
                var p5b = document.AddParagraph("Ellipse (Cyan)");
                var ellipse2 = p5b.AddShape(ShapeType.Ellipse, 70, 40, SixLabors.ImageSharp.Color.Cyan, SixLabors.ImageSharp.Color.Black, 1.25);
                ellipse2.Left = 160; ellipse2.Top = 120;

                var p5c = document.AddParagraph("RoundedRect (Yellow)");
                var rrect2 = p5c.AddShape(ShapeType.RoundedRectangle, 90, 45, SixLabors.ImageSharp.Color.Yellow, SixLabors.ImageSharp.Color.Black, 1);
                rrect2.Left = 300; rrect2.Top = 120;

                // DrawingML shapes (explicit fill/stroke and anchored absolute positioning)
                var p6 = document.AddParagraph("DrawingML Rectangle (theme)");
                var d1 = p6.AddShapeDrawing(ShapeType.Rectangle, 90, 30, 20, 220);
                d1.FillColor = SixLabors.ImageSharp.Color.LightSkyBlue;
                d1.StrokeColor = SixLabors.ImageSharp.Color.DarkBlue;
                d1.StrokeWeight = 1.5;

                var p7 = document.AddParagraph("DrawingML Triangle");
                var d2 = p7.AddShapeDrawing(ShapeType.Triangle, 70, 60, 140, 220);
                d2.FillColor = SixLabors.ImageSharp.Color.LightGreen;
                d2.StrokeColor = SixLabors.ImageSharp.Color.Black;

                var p8 = document.AddParagraph("DrawingML Diamond");
                var d3 = p8.AddShapeDrawing(ShapeType.Diamond, 70, 70, 240, 220);
                d3.FillColor = SixLabors.ImageSharp.Color.Coral;
                d3.StrokeColor = SixLabors.ImageSharp.Color.DarkRed;

                var p9 = document.AddParagraph("DrawingML Hexagon");
                var d4 = p9.AddShapeDrawing(ShapeType.Hexagon, 90, 60, 340, 220);
                d4.FillColor = SixLabors.ImageSharp.Color.Khaki;
                d4.StrokeColor = SixLabors.ImageSharp.Color.Olive;

                var p10 = document.AddParagraph("DrawingML Star5");
                var d5 = p10.AddShapeDrawing(ShapeType.Star5, 70, 70, 460, 220);
                d5.FillColor = SixLabors.ImageSharp.Color.Plum;
                d5.StrokeColor = SixLabors.ImageSharp.Color.Purple;

                document.Save(openWord);
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
