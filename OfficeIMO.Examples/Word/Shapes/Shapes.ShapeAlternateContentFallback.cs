using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
                Run run = word.MainDocumentPart!.Document!.Body!.Descendants<Run>().First(r => r.Descendants<Drawing>().Any());
                Drawing drawing = run.Descendants<Drawing>().First();
                Drawing fallbackDrawing = (Drawing)drawing.CloneNode(true);
                drawing.Remove();
                AlternateContentChoice choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(new Run(new Text("placeholder")));
                AlternateContentFallback fallback = new AlternateContentFallback();
                fallback.Append(fallbackDrawing);
                AlternateContent alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                run.Append(alt);
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
