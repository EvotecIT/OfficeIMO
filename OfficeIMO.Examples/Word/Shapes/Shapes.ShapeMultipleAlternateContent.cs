using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Path = System.IO.Path;

namespace OfficeIMO.Examples.Word {
    internal static partial class Shapes {
        internal static void Example_ShapeMultipleAlternateContent(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with shape and text box in separate AlternateContent elements");
            string filePath = Path.Combine(folderPath, "ShapeMultipleAlternateContent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                document.AddTextBox("Text");
                document.Save(false);
            }
            using (WordprocessingDocument word = WordprocessingDocument.Open(filePath, true)) {
                Body body = word.MainDocumentPart!.Document!.Body!;
                Run shapeRun = body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                Run textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                Drawing shapeDrawing = (Drawing)shapeRun.Descendants<Drawing>().First().CloneNode(true);
                Drawing textBoxDrawing = (Drawing)textBoxRun.Descendants<Drawing>().First().CloneNode(true);

                AlternateContent shapeAc = new AlternateContent();
                AlternateContentChoice shapeChoice = new AlternateContentChoice() { Requires = "wps" };
                shapeChoice.Append(shapeDrawing);
                shapeAc.Append(shapeChoice);

                AlternateContent textBoxAc = new AlternateContent();
                AlternateContentFallback textBoxFallback = new AlternateContentFallback();
                textBoxFallback.Append(textBoxDrawing);
                textBoxAc.Append(textBoxFallback);

                Run run = new Run();
                run.Append(shapeAc);
                run.Append(textBoxAc);

                shapeRun.Parent!.InsertBefore(run, shapeRun);
                shapeRun.Remove();
                textBoxRun.Remove();

                Document document = word.MainDocumentPart!.Document!;
                if (document.LookupNamespace("wps") == null) {
                    document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                }
                document.MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "wps" };
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Shapes count: {document.Shapes.Count}");
                Console.WriteLine($"TextBoxes count: {document.TextBoxes.Count}");
                document.Save(openWord);
            }
        }
    }
}
