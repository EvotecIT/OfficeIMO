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
                var rect = p1.AddShape(80, 40, OfficeIMO.OfficeColor.Green);

                // VML Ellipse with stroke
                var p2 = document.AddParagraph("Ellipse (Orange fill, Black stroke)");
                var ellipse = p2.AddShape(WordShapeType.Ellipse, 90, 50, OfficeIMO.OfficeColor.Orange, OfficeIMO.OfficeColor.Black, strokeWeightPt: 2);

                // VML Rounded Rectangle with arc size
                var p3 = document.AddParagraph("Rounded Rectangle (Blue)");
                var rrect = p3.AddShape(WordShapeType.RoundedRectangle, 100, 50, OfficeIMO.OfficeColor.Blue, OfficeIMO.OfficeColor.Black, strokeWeightPt: 1.5, arcSize: 0.4);

                // VML Line with stroke color/weight
                var p4 = document.AddParagraph("Line (Red, 3pt)");
                p4.AddLine(0, 0, 120, 0, OfficeIMO.OfficeColor.Red, 3);

                // VML Polygon (triangle)
                var p5 = document.AddParagraph("Polygon (Triangle, Purple)");
                var poly = WordShape.AddPolygon(p5, "0,0 60,0 30,50 0,0", OfficeIMO.OfficeColor.Purple, OfficeIMO.OfficeColor.Black);

                // Absolute positioning for VML shapes
                rect.Left = 20; rect.Top = 20; rect.Rotation = 5;
                ellipse.Left = 160; ellipse.Top = 20;
                rrect.Left = 300; rrect.Top = 20; rrect.Rotation = -10;
                poly.Left = 20; poly.Top = 120;

                // A second row of VML shapes for visibility (own paragraph)
                var p5b = document.AddParagraph("Ellipse (Cyan)");
                var ellipse2 = p5b.AddShape(WordShapeType.Ellipse, 70, 40, OfficeIMO.OfficeColor.Cyan, OfficeIMO.OfficeColor.Black, 1.25);
                ellipse2.Left = 160; ellipse2.Top = 120;

                var p5c = document.AddParagraph("RoundedRect (Yellow)");
                var rrect2 = p5c.AddShape(WordShapeType.RoundedRectangle, 90, 45, OfficeIMO.OfficeColor.Yellow, OfficeIMO.OfficeColor.Black, 1);
                rrect2.Left = 300; rrect2.Top = 120;

                // DrawingML shapes (explicit fill/stroke and anchored absolute positioning)
                var p6 = document.AddParagraph("DrawingML Rectangle (theme)");
                var d1 = p6.AddShapeDrawing(WordShapeType.Rectangle, 90, 30, 20, 220);
                d1.FillColor = OfficeIMO.OfficeColor.LightSkyBlue;
                d1.StrokeColor = OfficeIMO.OfficeColor.DarkBlue;
                d1.StrokeWeight = 1.5;

                var p7 = document.AddParagraph("DrawingML Triangle");
                var d2 = p7.AddShapeDrawing(WordShapeType.Triangle, 70, 60, 140, 220);
                d2.FillColor = OfficeIMO.OfficeColor.LightGreen;
                d2.StrokeColor = OfficeIMO.OfficeColor.Black;

                var p8 = document.AddParagraph("DrawingML Diamond");
                var d3 = p8.AddShapeDrawing(WordShapeType.Diamond, 70, 70, 240, 220);
                d3.FillColor = OfficeIMO.OfficeColor.Coral;
                d3.StrokeColor = OfficeIMO.OfficeColor.DarkRed;

                var p9 = document.AddParagraph("DrawingML Hexagon");
                var d4 = p9.AddShapeDrawing(WordShapeType.Hexagon, 90, 60, 340, 220);
                d4.FillColor = OfficeIMO.OfficeColor.Khaki;
                d4.StrokeColor = OfficeIMO.OfficeColor.Olive;

                var p10 = document.AddParagraph("DrawingML Star5");
                var d5 = p10.AddShapeDrawing(WordShapeType.Star5, 70, 70, 460, 220);
                d5.FillColor = OfficeIMO.OfficeColor.Plum;
                d5.StrokeColor = OfficeIMO.OfficeColor.Purple;

                document.Save();
                if (openWord) document.OpenInApplication();
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
