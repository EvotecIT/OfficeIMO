using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using V = DocumentFormat.OpenXml.Vml;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;
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
                doc.Save();
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var mainPart = wDoc.MainDocumentPart;
                Assert.NotNull(mainPart);
                var document = mainPart!.Document;
                Assert.NotNull(document);
                var body = document!.Body;
                Assert.NotNull(body);
                var oval = body!.Descendants<V.Oval>().First();
                var textBox = new V.TextBox();
                textBox.Append(new TextBoxContent(new Paragraph(new Run(new Text("Text")))));
                oval.Append(textBox);
                document.Save();
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
                doc.Save();
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var mainPart = wDoc.MainDocumentPart;
                Assert.NotNull(mainPart);
                var document = mainPart!.Document;
                Assert.NotNull(document);
                var body = document!.Body;
                Assert.NotNull(body);
                var run = body!.Descendants<Run>().First(r => r.Descendants<WordDrawing>().Any());
                var drawing = run.Descendants<WordDrawing>().First();
                drawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(drawing);
                var alt = new AlternateContent();
                alt.Append(choice);
                run.Append(alt);
                document.Save();
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
                doc.Save();
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var mainPart = wDoc.MainDocumentPart;
                Assert.NotNull(mainPart);
                var document = mainPart!.Document;
                Assert.NotNull(document);
                var body = document!.Body;
                Assert.NotNull(body);
                var run = body!.Descendants<Run>().First(r => r.Descendants<WordDrawing>().Any());
                var drawing = run.Descendants<WordDrawing>().First();
                var fallbackDrawing = (WordDrawing)drawing.CloneNode(true);
                drawing.Remove();
                var choice = new AlternateContentChoice() { Requires = "wps" };
                choice.Append(new Run(new Text("placeholder")));
                var fallback = new AlternateContentFallback();
                fallback.Append(fallbackDrawing);
                var alt = new AlternateContent();
                alt.Append(choice);
                alt.Append(fallback);
                run.Append(alt);
                document.Save();
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
                doc.Save();
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var mainPart = wDoc.MainDocumentPart;
                Assert.NotNull(mainPart);
                var document = mainPart!.Document;
                Assert.NotNull(document);
                var body = document!.Body;
                Assert.NotNull(body);
                var shapeRun = body!.Descendants<Run>().First(r => r.Descendants<WordDrawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                var textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                var shapeDrawing = shapeRun.Descendants<WordDrawing>().First();
                var textBoxDrawing = textBoxRun.Descendants<WordDrawing>().First();
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
                doc.Save();
            }
            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(filePath, true)) {
                var mainPart = wDoc.MainDocumentPart;
                Assert.NotNull(mainPart);
                var document = mainPart!.Document;
                Assert.NotNull(document);
                var body = document!.Body;
                Assert.NotNull(body);
                var shapeRun = body!.Descendants<Run>().First(r => r.Descendants<WordDrawing>().Any() && !r.Descendants<Wps.TextBoxInfo2>().Any());
                var textBoxRun = body.Descendants<Run>().First(r => r.Descendants<Wps.TextBoxInfo2>().Any());
                var shapeDrawing = (WordDrawing)shapeRun.Descendants<WordDrawing>().First().CloneNode(true);
                var textBoxDrawing = (WordDrawing)textBoxRun.Descendants<WordDrawing>().First().CloneNode(true);

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

                var parent = shapeRun.Parent;
                Assert.NotNull(parent);
                parent!.InsertBefore(run, shapeRun);
                shapeRun.Remove();
                textBoxRun.Remove();

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
