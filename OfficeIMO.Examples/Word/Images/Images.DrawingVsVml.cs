using System;
using System.IO;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_DrawingVsVml(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating DrawingML and VML sample documents");

            string assets = Path.Combine(Directory.GetCurrentDirectory(), "Assets");
            string img = Path.Combine(assets, "OfficeIMO.png");

            string drawingFile = Path.Combine(folderPath, "DrawingVsVml.Drawing.docx");
            using (WordDocument doc = WordDocument.Create(drawingFile)) {
                doc.AddParagraph().AddImage(img);
                doc.AddShapeDrawing(ShapeType.Ellipse, 40, 40);
                doc.AddShapeDrawing(ShapeType.Rectangle, 60, 30);
                doc.AddShapeDrawing(ShapeType.RoundedRectangle, 50, 30);
                doc.AddTextBox("Text");
                doc.Save(false);
            }

            string vmlFile = Path.Combine(folderPath, "DrawingVsVml.Vml.docx");
            using (WordDocument doc = WordDocument.Create(vmlFile)) {
                doc.AddImageVml(img);
                doc.AddShape(ShapeType.Ellipse, 40, 40, Color.Red, Color.Blue);
                doc.AddShape(ShapeType.Rectangle, 60, 30, Color.Green, Color.Black);
                doc.AddShape(ShapeType.RoundedRectangle, 50, 30, Color.Yellow, Color.Black, 1, arcSize: 0.3);
                doc.AddTextBoxVml("Text");
                doc.Save(false);
            }

            using (WordDocument doc = WordDocument.Load(drawingFile)) {
                Console.WriteLine($"DrawingML -> Images: {doc.Images.Count}, Shapes: {doc.Shapes.Count}, TextBoxes: {doc.TextBoxes.Count}");
            }

            using (WordDocument doc = WordDocument.Load(vmlFile)) {
                Console.WriteLine($"VML -> Images: {doc.Images.Count}, Shapes: {doc.Shapes.Count}, TextBoxes: {doc.TextBoxes.Count}");
            }

            if (openWord) {
                Console.WriteLine("OpenWord flag is set but opening is not implemented in this environment.");
            }
        }
    }
}

