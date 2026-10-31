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
                var rect = p1.AddShape(OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeColor.Green);

                // VML Ellipse with stroke
                var p2 = document.AddParagraph("Ellipse (Orange fill, Black stroke)");
                var ellipse = p2.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeColor.Orange, OfficeIMO.OfficeColor.Black, strokeWeight: OfficeIMO.OfficeLength.Points(2));

                // VML Rounded Rectangle with arc size
                var p3 = document.AddParagraph("Rounded Rectangle (Blue)");
                var rrect = p3.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeColor.Blue, OfficeIMO.OfficeColor.Black, strokeWeight: OfficeIMO.OfficeLength.Points(1.5), arcSize: 0.4);

                // VML Line with stroke color/weight
                var p4 = document.AddParagraph("Line (Red, 3pt)");
                p4.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Red, OfficeIMO.OfficeLength.Points(3));

                // VML Polygon (triangle)
                var p5 = document.AddParagraph("Polygon (Triangle, Purple)");
                var poly = WordShape.AddPolygon(p5, "0,0 60,0 30,50 0,0", OfficeIMO.OfficeColor.Purple, OfficeIMO.OfficeColor.Black);

                // Absolute positioning for VML shapes
                rect.Left = OfficeIMO.OfficeLength.Points(20); rect.Top = OfficeIMO.OfficeLength.Points(20); rect.Rotation = 5;
                ellipse.Left = OfficeIMO.OfficeLength.Points(160); ellipse.Top = OfficeIMO.OfficeLength.Points(20);
                rrect.Left = OfficeIMO.OfficeLength.Points(300); rrect.Top = OfficeIMO.OfficeLength.Points(20); rrect.Rotation = -10;
                poly.Left = OfficeIMO.OfficeLength.Points(20); poly.Top = OfficeIMO.OfficeLength.Points(120);

                // A second row of VML shapes for visibility (own paragraph)
                var p5b = document.AddParagraph("Ellipse (Cyan)");
                var ellipse2 = p5b.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeColor.Cyan, OfficeIMO.OfficeColor.Black, OfficeIMO.OfficeLength.Points(1.25));
                ellipse2.Left = OfficeIMO.OfficeLength.Points(160); ellipse2.Top = OfficeIMO.OfficeLength.Points(120);

                var p5c = document.AddParagraph("RoundedRect (Yellow)");
                var rrect2 = p5c.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(45), OfficeIMO.OfficeColor.Yellow, OfficeIMO.OfficeColor.Black, OfficeIMO.OfficeLength.Points(1));
                rrect2.Left = OfficeIMO.OfficeLength.Points(300); rrect2.Top = OfficeIMO.OfficeLength.Points(120);

                // DrawingML shapes (explicit fill/stroke and anchored absolute positioning)
                var p6 = document.AddParagraph("DrawingML Rectangle (theme)");
                var d1 = p6.AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(220));
                d1.FillColor = OfficeIMO.OfficeColor.LightSkyBlue;
                d1.StrokeColor = OfficeIMO.OfficeColor.DarkBlue;
                d1.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p7 = document.AddParagraph("DrawingML Triangle");
                var d2 = p7.AddShapeDrawing(WordShapeType.Triangle, OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(140), OfficeIMO.OfficeLength.Points(220));
                d2.FillColor = OfficeIMO.OfficeColor.LightGreen;
                d2.StrokeColor = OfficeIMO.OfficeColor.Black;

                var p8 = document.AddParagraph("DrawingML Diamond");
                var d3 = p8.AddShapeDrawing(WordShapeType.Diamond, OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(240), OfficeIMO.OfficeLength.Points(220));
                d3.FillColor = OfficeIMO.OfficeColor.Coral;
                d3.StrokeColor = OfficeIMO.OfficeColor.DarkRed;

                var p9 = document.AddParagraph("DrawingML Hexagon");
                var d4 = p9.AddShapeDrawing(WordShapeType.Hexagon, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(340), OfficeIMO.OfficeLength.Points(220));
                d4.FillColor = OfficeIMO.OfficeColor.Khaki;
                d4.StrokeColor = OfficeIMO.OfficeColor.Olive;

                var p10 = document.AddParagraph("DrawingML Star5");
                var d5 = p10.AddShapeDrawing(WordShapeType.Star5, OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(460), OfficeIMO.OfficeLength.Points(220));
                d5.FillColor = OfficeIMO.OfficeColor.Plum;
                d5.StrokeColor = OfficeIMO.OfficeColor.Purple;

                document.Save();
                if (openWord) document.OpenInApplication();
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}
