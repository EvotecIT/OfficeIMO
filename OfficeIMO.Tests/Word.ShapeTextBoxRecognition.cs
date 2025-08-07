using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Xunit;
using Color = SixLabors.ImageSharp.Color;
using Path = System.IO.Path;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_VmlEllipseWithTextRecognizedAsTextBoxOnly() {
            string filePath = Path.Combine(_directoryWithFiles, "EllipseShapeWithText.docx");
            using (WordDocument doc = WordDocument.Create(filePath)) {
                doc.AddShape(ShapeType.Ellipse, 40, 40, Color.Red, Color.Blue);
                doc.Save(false);
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var oval = wDoc.MainDocumentPart.Document.Body.Descendants<V.Oval>().First();
                var textBox = new V.TextBox();
                textBox.Append(new TextBoxContent(new Paragraph(new Run(new Text("Text")))));
                oval.Append(textBox);
                wDoc.MainDocumentPart.Document.Save();
            }
            using (WordDocument doc = WordDocument.Load(filePath)) {
                Assert.Empty(doc.Shapes);
                Assert.Single(doc.TextBoxes);
            }
        }

        [Fact]
        public void Test_AlternateContentShapeNotTreatedAsTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "ShapeWrappedInAlternateContent.docx");
            using (WordDocument doc = WordDocument.Create(filePath)) {
                doc.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                doc.Save(false);
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var run = wDoc.MainDocumentPart.Document.Body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any());
                var drawing = run.Descendants<Drawing>().First();
                drawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(drawing);
                var alt = new AlternateContent();
                alt.Append(choice);
                run.Append(alt);
                wDoc.MainDocumentPart.Document.Save();
            }
            using (WordDocument doc = WordDocument.Load(filePath)) {
                Assert.Single(doc.Shapes);
                Assert.Empty(doc.TextBoxes);
            }
        }

        [Fact]
        public void Test_AlternateContentFallbackShapeDetected() {
            string filePath = Path.Combine(_directoryWithFiles, "ShapeInAlternateContentFallback.docx");
            using (WordDocument doc = WordDocument.Create(filePath)) {
                doc.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                doc.Save(false);
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var run = wDoc.MainDocumentPart.Document.Body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any());
                var drawing = run.Descendants<Drawing>().First();
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
                wDoc.MainDocumentPart.Document.Save();
            }
            using (WordDocument doc = WordDocument.Load(filePath)) {
                Assert.Single(doc.Shapes);
                Assert.Empty(doc.TextBoxes);
            }
        }

        [Fact]
        public void Test_AlternateContentChoiceShapeFallbackTextBox() {
            string filePath = Path.Combine(_directoryWithFiles, "ShapeChoiceFallbackTextBox.docx");
            using (WordDocument doc = WordDocument.Create(filePath)) {
                doc.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                doc.AddTextBox("Text");
                doc.Save(false);
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var body = wDoc.MainDocumentPart.Document.Body;
                var shapeRun = body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                var textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                var shapeDrawing = shapeRun.Descendants<Drawing>().First();
                var textBoxDrawing = textBoxRun.Descendants<Drawing>().First();
                shapeDrawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(shapeDrawing);
                var fallback = new AlternateContentFallback();
                fallback.Append((Drawing)textBoxDrawing.CloneNode(true));
                var alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                shapeRun.Append(alt);
                textBoxRun.Remove();
                var document = wDoc.MainDocumentPart.Document;
                if (document.LookupNamespace("wps") == null) {
                    document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                }
                document.MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "wps" };
                document.Save();
            }
            using (WordDocument doc = WordDocument.Load(filePath)) {
                Assert.Single(doc.Shapes);
                Assert.Empty(doc.TextBoxes);
            }
        }

        [Fact]
        public void Test_MultipleAlternateContentTextBoxPreferred() {
            string filePath = Path.Combine(_directoryWithFiles, "MultipleAlternateContentTextBoxPreferred.docx");
            using (WordDocument doc = WordDocument.Create(filePath)) {
                doc.AddShapeDrawing(ShapeType.Rectangle, 40, 40);
                doc.AddTextBox("Text");
                doc.Save(false);
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var body = wDoc.MainDocumentPart.Document.Body;
                var shapeRun = body.Descendants<Run>().First(r => r.Descendants<Drawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                var textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                var shapeDrawing = (Drawing)shapeRun.Descendants<Drawing>().First().CloneNode(true);
                var textBoxDrawing = (Drawing)textBoxRun.Descendants<Drawing>().First().CloneNode(true);

                var shapeAc = new AlternateContent();
                var shapeChoice = new AlternateContentChoice() { Requires = "wps" };
                shapeChoice.Append(shapeDrawing);
                shapeAc.Append(shapeChoice);

                var textBoxAc = new AlternateContent();
                var textBoxFallback = new AlternateContentFallback();
                textBoxFallback.Append(textBoxDrawing);
                textBoxAc.Append(textBoxFallback);

                var run = new Run();
                run.Append(shapeAc);
                run.Append(textBoxAc);

                shapeRun.Parent.InsertBefore(run, shapeRun);
                shapeRun.Remove();
                textBoxRun.Remove();

                var document = wDoc.MainDocumentPart.Document;
                if (document.LookupNamespace("wps") == null) {
                    document.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                }
                document.MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "wps" };
                document.Save();
            }
            using (WordDocument doc = WordDocument.Load(filePath)) {
                Assert.Empty(doc.Shapes);
                Assert.Single(doc.TextBoxes);
            }
        }

    }
}
