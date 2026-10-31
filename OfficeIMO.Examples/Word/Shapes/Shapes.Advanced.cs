using System;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapesAdvanced(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with advanced shape features");
            string filePath = System.IO.Path.Combine(folderPath, "DocumentWithAdvancedShapes.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                // VML rectangle
                var p1 = document.AddParagraph("VML Rectangle (Aqua fill, Navy stroke)");
                var rect = p1.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(40), Color.Aqua, Color.Navy, OfficeIMO.OfficeLength.Points(2));
                rect.Left = OfficeIMO.OfficeLength.Points(20); rect.Top = OfficeIMO.OfficeLength.Points(20);

                // VML ellipse rotated
                var p2 = document.AddParagraph("VML Ellipse rotated 25°");
                var ell = p2.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(40), Color.Gold, Color.Black, OfficeIMO.OfficeLength.Points(1.25));
                ell.Rotation = 25; ell.Left = OfficeIMO.OfficeLength.Points(160); ell.Top = OfficeIMO.OfficeLength.Points(20);

                // VML rounded rectangle with larger arc
                var p3 = document.AddParagraph("VML RoundedRect arc=0.6");
                var rrect = p3.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(50), Color.LightGreen, Color.DarkGreen, OfficeIMO.OfficeLength.Points(1), arcSize: 0.6);
                rrect.Left = OfficeIMO.OfficeLength.Points(300); rrect.Top = OfficeIMO.OfficeLength.Points(20);

                // VML polygon triangle
                var p4 = document.AddParagraph("VML Polygon (triangle)");
                var poly = WordShape.AddPolygon(p4, "0,0 60,0 30,50 0,0", Color.MediumPurple, Color.Black);
                poly.Left = OfficeIMO.OfficeLength.Points(20); poly.Top = OfficeIMO.OfficeLength.Points(120);

                // VML line
                var p5 = document.AddParagraph("VML Line 120pt (Red)");
                p5.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(120), OfficeIMO.OfficeLength.Points(0), Color.Red, OfficeIMO.OfficeLength.Points(2.5));

                // DrawingML rectangle with explicit fill/stroke (anchored absolute)
                var p6 = document.AddParagraph("DrawingML Rectangle with explicit fill/stroke");
                var dml = p6.AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(240));
                dml.FillColor = Color.CornflowerBlue; // ensure visible fill
                dml.StrokeColor = Color.SaddleBrown;  // ensure visible outline
                dml.StrokeWeight = OfficeIMO.OfficeLength.Points(2);

                // More DrawingML shapes (each on its own paragraph)
                var p7 = document.AddParagraph("DrawingML RightArrow");
                var d7 = p7.AddShapeDrawing(WordShapeType.RightArrow, OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(140), OfficeIMO.OfficeLength.Points(240));
                d7.FillColor = Color.LightSkyBlue;
                d7.StrokeColor = Color.DarkBlue;
                d7.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p8 = document.AddParagraph("DrawingML LeftArrow");
                var d8 = p8.AddShapeDrawing(WordShapeType.LeftArrow, OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(260), OfficeIMO.OfficeLength.Points(240));
                d8.FillColor = Color.Khaki;
                d8.StrokeColor = Color.Olive;
                d8.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p9 = document.AddParagraph("DrawingML UpArrow");
                var d9 = p9.AddShapeDrawing(WordShapeType.UpArrow, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(380), OfficeIMO.OfficeLength.Points(220));
                d9.FillColor = Color.LightGreen;
                d9.StrokeColor = Color.DarkGreen;
                d9.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p10 = document.AddParagraph("DrawingML DownArrow");
                var d10 = p10.AddShapeDrawing(WordShapeType.DownArrow, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(460), OfficeIMO.OfficeLength.Points(220));
                d10.FillColor = Color.Salmon;
                d10.StrokeColor = Color.Maroon;
                d10.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p11 = document.AddParagraph("DrawingML Heart");
                var d11 = p11.AddShapeDrawing(WordShapeType.Heart, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(340));
                d11.FillColor = Color.Pink;
                d11.StrokeColor = Color.HotPink;
                d11.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p12 = document.AddParagraph("DrawingML Cloud");
                var d12 = p12.AddShapeDrawing(WordShapeType.Cloud, OfficeIMO.OfficeLength.Points(110), OfficeIMO.OfficeLength.Points(70), OfficeIMO.OfficeLength.Points(140), OfficeIMO.OfficeLength.Points(340));
                d12.FillColor = Color.LightGray;
                d12.StrokeColor = Color.DimGray;
                d12.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p13 = document.AddParagraph("DrawingML Donut");
                var d13 = p13.AddShapeDrawing(WordShapeType.Donut, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(280), OfficeIMO.OfficeLength.Points(330));
                d13.FillColor = Color.Gold;
                d13.StrokeColor = Color.DarkGoldenrod;
                d13.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p14 = document.AddParagraph("DrawingML Can");
                var d14 = p14.AddShapeDrawing(WordShapeType.Can, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(320));
                d14.FillColor = Color.LightSteelBlue;
                d14.StrokeColor = Color.SteelBlue;
                d14.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);

                var p15 = document.AddParagraph("DrawingML Cube");
                var d15 = p15.AddShapeDrawing(WordShapeType.Cube, OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(90), OfficeIMO.OfficeLength.Points(500), OfficeIMO.OfficeLength.Points(330));
                d15.FillColor = Color.Plum;
                d15.StrokeColor = Color.Purple;
                d15.StrokeWeight = OfficeIMO.OfficeLength.Points(1.5);
                document.Save();
                if (openWord) document.OpenInApplication();
                OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
            }
        }
    }
}

