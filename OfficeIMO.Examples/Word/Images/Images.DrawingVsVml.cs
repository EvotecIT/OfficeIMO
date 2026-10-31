using System;
using System.IO;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word {
    internal static partial class Images {
        internal static void Example_DrawingVsVml(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating DrawingML and VML sample documents");

            string assets = Path.Combine(Directory.GetCurrentDirectory(), "Assets");
            string img = Path.Combine(assets, "OfficeIMO.png");

            string drawingFile = Path.Combine(folderPath, "DrawingVsVml.Drawing.docx");
            using (WordDocument doc = WordDocument.Create(drawingFile)) {
                doc.AddParagraph().AddImage(img);
                doc.AddShapeDrawing(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(40));
                doc.AddShapeDrawing(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(30));
                doc.AddShapeDrawing(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(30));
                doc.AddTextBox("Text");
                doc.Save();
            }

            string vmlFile = Path.Combine(folderPath, "DrawingVsVml.Vml.docx");
            using (WordDocument doc = WordDocument.Create(vmlFile)) {
                doc.AddImage(img,options:new OfficeIMO.Word.WordImageOptions { Vml = true });
                doc.AddShape(WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(40), Color.Red, Color.Blue);
                doc.AddShape(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(60), OfficeIMO.OfficeLength.Points(30), Color.Green, Color.Black);
                doc.AddShape(WordShapeType.RoundedRectangle, OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(30), Color.Yellow, Color.Black, OfficeIMO.OfficeLength.Points(1), arcSize: 0.3);
                doc.AddTextBoxVml("Text");
                doc.Save();
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

