using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using DocumentFormat.OpenXml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

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
                var body = Guard.NotNull(word.MainDocumentPart?.Document?.Body, "Document body must exist.");
                var shapeRun = body.Descendants<Run>().FirstOrDefault(r => r.Descendants<WordDrawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any())
                    ?? throw new InvalidOperationException("Run containing a shape drawing was not found.");
                var textBoxRun = body.Descendants<Run>().FirstOrDefault(r => r.Descendants<Wps.TextBoxInfo2>().Any())
                    ?? throw new InvalidOperationException("Run containing a textbox drawing was not found.");
                var shapeDrawing = shapeRun.Descendants<WordDrawing>().FirstOrDefault()
                    ?? throw new InvalidOperationException("Shape drawing was not found.");
                var textBoxDrawing = textBoxRun.Descendants<WordDrawing>().FirstOrDefault()
                    ?? throw new InvalidOperationException("Textbox drawing was not found.");
                shapeDrawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(shapeDrawing);
                var fallback = new AlternateContentFallback();
                fallback.Append((WordDrawing)textBoxDrawing.CloneNode(true));
                var alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                shapeRun.Append(alt);
                textBoxRun.Remove();
                Guard.NotNull(word.MainDocumentPart?.Document, "Main document part must contain a document.").Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Shapes count: {document.Shapes.Count}");
                Console.WriteLine($"TextBoxes count: {document.TextBoxes.Count}");
                document.Save(openWord);
            }
        }
    }
}
