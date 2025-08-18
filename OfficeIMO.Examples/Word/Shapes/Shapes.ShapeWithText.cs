using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using V = DocumentFormat.OpenXml.Vml;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapeWithTextRecognizedAsTextBox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with ellipse shape that contains text");
            string filePath = System.IO.Path.Combine(folderPath, "ShapeWithText.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddShape(ShapeType.Ellipse, 40, 40, Color.Red, Color.Blue);
                document.Save(false);
            }
            using (WordprocessingDocument word = WordprocessingDocument.Open(filePath, true)) {
                V.Oval oval = word.MainDocumentPart!.Document!.Body!.Descendants<V.Oval>().First();
                V.TextBox textBox = new V.TextBox();
                textBox.Append(new TextBoxContent(new Paragraph(new Run(new Text("Text")))));
                oval.Append(textBox);
                word.MainDocumentPart!.Document!.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Shapes count: {document.Shapes.Count}");
                Console.WriteLine($"TextBoxes count: {document.TextBoxes.Count}");
                document.Save(openWord);
            }
        }
    }
}
