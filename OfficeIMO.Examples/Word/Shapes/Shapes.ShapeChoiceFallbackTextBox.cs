using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using DocumentFormat.OpenXml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapeChoiceFallbackTextBox(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shape in choice and text box in fallback");
            string filePath = Path.Combine(folderPath, "ShapeChoiceFallbackTextBox.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                document.AddTextBox("Text");
                document.Save(false);
            }
            using (WordprocessingDocument word = WordprocessingDocument.Open(filePath, true)) {
                Body body = word.MainDocumentPart!.Document!.Body!;
                Run shapeRun = body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                Run textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                Drawing shapeDrawing = shapeRun.Descendants<Drawing>().First();
                Drawing textBoxDrawing = textBoxRun.Descendants<Drawing>().First();
                shapeDrawing.Remove();
                AlternateContentChoice choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(shapeDrawing);
                AlternateContentFallback fallback = new AlternateContentFallback();
                fallback.Append((Drawing)textBoxDrawing.CloneNode(true));
                AlternateContent alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                shapeRun.Append(alt);
                textBoxRun.Remove();
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
