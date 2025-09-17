using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Examples.Utils;
using OfficeIMO.Word;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapeInAlternateContentFallback(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shape inside AlternateContent fallback");
            string filePath = System.IO.Path.Combine(folderPath, "ShapeInAlternateContentFallback.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                document.Save(false);
            }
            using (WordprocessingDocument word = WordprocessingDocument.Open(filePath, true)) {
                var body = Guard.NotNull(word.MainDocumentPart?.Document?.Body, "Document body must exist.");
                var run = body.Descendants<Run>().FirstOrDefault(r => r.Descendants<Drawing>().Any())
                    ?? throw new InvalidOperationException("No run containing a drawing was found.");
                var drawing = run.Descendants<Drawing>().FirstOrDefault()
                    ?? throw new InvalidOperationException("Expected drawing element to be present.");
                var fallbackDrawing = (Drawing)drawing.CloneNode(true);
                drawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(new Run(new Text("placeholder")));
                var fallback = new AlternateContentFallback();
                fallback.Append(fallbackDrawing);
                var alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                run.Append(alt);
                var mainDocument = Guard.NotNull(word.MainDocumentPart?.Document, "Main document part must expose a document.");
                mainDocument.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Shapes count: {document.Shapes.Count}");
                Console.WriteLine($"TextBoxes count: {document.TextBoxes.Count}");
                document.Save(openWord);
            }
        }
    }
}
