using System;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapesAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced shape features");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithAdvancedShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // VML rectangle
                var p1 = document.AddParagraph("VML Rectangle (Aqua fill, Navy stroke)");
                var rect = p1.AddShape(ShapeType.Rectangle, 90, 40, Color.Aqua, Color.Navy, 2);
                rect.Left = 20; rect.Top = 20;

                // VML ellipse rotated
                var p2 = document.AddParagraph("VML Ellipse rotated 25°");
                var ell = p2.AddShape(ShapeType.Ellipse, 80, 40, Color.Gold, Color.Black, 1.25);
                ell.Rotation = 25; ell.Left = 160; ell.Top = 20;

                // VML rounded rectangle with larger arc
                var p3 = document.AddParagraph("VML RoundedRect arc=0.6");
                var rrect = p3.AddShape(ShapeType.RoundedRectangle, 100, 50, Color.LightGreen, Color.DarkGreen, 1, arcSize: 0.6);
                rrect.Left = 300; rrect.Top = 20;

                // VML polygon triangle
                var p4 = document.AddParagraph("VML Polygon (triangle)");
                var poly = WordShape.AddPolygon(p4, "0,0 60,0 30,50 0,0", Color.MediumPurple, Color.Black);
                poly.Left = 20; poly.Top = 120;

                // VML line
                var p5 = document.AddParagraph("VML Line 120pt (Red)");
                p5.AddLine(0, 0, 120, 0, Color.Red, 2.5);

                // DrawingML rectangle with explicit fill/stroke (anchored absolute)
                var p6 = document.AddParagraph("DrawingML Rectangle with explicit fill/stroke");
                var dml = p6.AddShapeDrawing(ShapeType.Rectangle, 90, 30, 20, 240);
                dml.FillColor = Color.CornflowerBlue; // ensure visible fill
                dml.StrokeColor = Color.SaddleBrown;  // ensure visible outline
                dml.StrokeWeight = 2;

                // More DrawingML shapes (each on its own paragraph)
                var p7 = document.AddParagraph("DrawingML RightArrow");
                var d7 = p7.AddShapeDrawing(ShapeType.RightArrow, 100, 40, 140, 240);
                d7.FillColor = Color.LightSkyBlue;
                d7.StrokeColor = Color.DarkBlue;
                d7.StrokeWeight = 1.5;

                var p8 = document.AddParagraph("DrawingML LeftArrow");
                var d8 = p8.AddShapeDrawing(ShapeType.LeftArrow, 100, 40, 260, 240);
                d8.FillColor = Color.Khaki;
                d8.StrokeColor = Color.Olive;
                d8.StrokeWeight = 1.5;

                var p9 = document.AddParagraph("DrawingML UpArrow");
                var d9 = p9.AddShapeDrawing(ShapeType.UpArrow, 60, 90, 380, 220);
                d9.FillColor = Color.LightGreen;
                d9.StrokeColor = Color.DarkGreen;
                d9.StrokeWeight = 1.5;

                var p10 = document.AddParagraph("DrawingML DownArrow");
                var d10 = p10.AddShapeDrawing(ShapeType.DownArrow, 60, 90, 460, 220);
                d10.FillColor = Color.Salmon;
                d10.StrokeColor = Color.Maroon;
                d10.StrokeWeight = 1.5;

                var p11 = document.AddParagraph("DrawingML Heart");
                var d11 = p11.AddShapeDrawing(ShapeType.Heart, 80, 70, 20, 340);
                d11.FillColor = Color.Pink;
                d11.StrokeColor = Color.HotPink;
                d11.StrokeWeight = 1.5;

                var p12 = document.AddParagraph("DrawingML Cloud");
                var d12 = p12.AddShapeDrawing(ShapeType.Cloud, 110, 70, 140, 340);
                d12.FillColor = Color.LightGray;
                d12.StrokeColor = Color.DimGray;
                d12.StrokeWeight = 1.5;

                var p13 = document.AddParagraph("DrawingML Donut");
                var d13 = p13.AddShapeDrawing(ShapeType.Donut, 90, 90, 280, 330);
                d13.FillColor = Color.Gold;
                d13.StrokeColor = Color.DarkGoldenrod;
                d13.StrokeWeight = 1.5;

                var p14 = document.AddParagraph("DrawingML Can");
                var d14 = p14.AddShapeDrawing(ShapeType.Can, 80, 100, 400, 320);
                d14.FillColor = Color.LightSteelBlue;
                d14.StrokeColor = Color.SteelBlue;
                d14.StrokeWeight = 1.5;

                var p15 = document.AddParagraph("DrawingML Cube");
                var d15 = p15.AddShapeDrawing(ShapeType.Cube, 90, 90, 500, 330);
                d15.FillColor = Color.Plum;
                d15.StrokeColor = Color.Purple;
                d15.StrokeWeight = 1.5;
                document.Save(openWord);
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}

