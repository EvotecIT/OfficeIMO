using System.IO;
using OfficeIMO.Word;
using SixLabors.ImageSharp;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithShapes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with shape");
                var shape = paragraph.AddShape(100, 50, Color.Red);
                shape.Title = "Rectangle";
                shape.Description = "My rectangle";
                shape.Hidden = false;
                shape.Stroked = true;
                shape.StrokeColor = Color.Blue;
                shape.StrokeWeight = 2;
                shape.Left = 10;
                shape.Top = 20;
                shape.Rotation = 45;
                shape.Width = 120;
                shape.Height = 60;

                WordShape.AddEllipse(paragraph, 80, 40, Color.Lime);
                WordShape.AddPolygon(paragraph, "0,0 50,0 50,50 0,50", Color.White, Color.Blue);
                WordShape.AddLine(paragraph, 0, 60, 100, 60, Color.Red, 2);

                Assert.True(document.Paragraphs.Count == 1);
                Assert.NotNull(paragraph.Shape);
                Assert.Equal("Rectangle", shape.Title);
                Assert.Equal("My rectangle", shape.Description);
                Assert.False(shape.Hidden!.Value);
                Assert.True(shape.Stroked!.Value);
                Assert.Equal(Color.Blue.ToHexColor(), shape.StrokeColorHex);
                Assert.Equal(2d, shape.StrokeWeight!.Value, 1);
                Assert.Equal(120d, shape.Width, 1);
                Assert.Equal(60d, shape.Height, 1);
                Assert.Equal(10d, shape.Left!.Value, 1);
                Assert.Equal(20d, shape.Top!.Value, 1);
                Assert.Equal(45d, shape.Rotation!.Value, 1);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithShapes.docx"))) {
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.Equal("Rectangle", document.Paragraphs[0].Shape.Title);
                Assert.Equal("My rectangle", document.Paragraphs[0].Shape.Description);
                Assert.False(document.Paragraphs[0].Shape.Hidden!.Value);
                Assert.True(document.Paragraphs[0].Shape.Stroked!.Value);
                Assert.Equal(Color.Blue.ToHexColor(), document.Paragraphs[0].Shape.StrokeColorHex);
                Assert.Equal(2d, document.Paragraphs[0].Shape.StrokeWeight!.Value, 1);
                Assert.Equal(120d, document.Paragraphs[0].Shape.Width, 1);
                Assert.Equal(60d, document.Paragraphs[0].Shape.Height, 1);
                Assert.Equal(10d, document.Paragraphs[0].Shape.Left!.Value, 1);
                Assert.Equal(20d, document.Paragraphs[0].Shape.Top!.Value, 1);
                Assert.Equal(45d, document.Paragraphs[0].Shape.Rotation!.Value, 1);
            }
        }

        [Fact]
        public void Test_AddShapeFromDocument() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentAddShape.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var shape = document.AddShape(ShapeType.Rectangle, 80, 40, Color.Lime, Color.Black, 2);
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.Equal(Color.Lime.ToHexColor(), shape.FillColorHex);
                Assert.Equal(Color.Black.ToHexColor(), shape.StrokeColorHex);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.Equal(Color.Lime.ToHexColor(), document.Paragraphs[0].Shape.FillColorHex);
                Assert.Equal(Color.Black.ToHexColor(), document.Paragraphs[0].Shape.StrokeColorHex);
            }
        }

        [Fact]
        public void Test_AddShapeOnParagraphEnum() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphAddShape.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                var shape = paragraph.AddShape(ShapeType.Ellipse, 60, 30, Color.Aqua, Color.Red, 1.5);
                Assert.True(paragraph.IsShape);
                Assert.Equal(Color.Aqua.ToHexColor(), shape.FillColorHex);
                Assert.Equal(Color.Red.ToHexColor(), shape.StrokeColorHex);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.Equal(Color.Aqua.ToHexColor(), document.Paragraphs[0].Shape.FillColorHex);
                Assert.Equal(Color.Red.ToHexColor(), document.Paragraphs[0].Shape.StrokeColorHex);
            }
        }

        [Fact]
        public void Test_ShapeCollectionsAndRemoval() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentShapesCollections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var rect = document.AddShape(ShapeType.Rectangle, 50, 20);
                var ellipse = document.AddShape(ShapeType.Ellipse, 40, 40, Color.Green, Color.Blue);
                ellipse.FillColor = Color.Yellow;

                Assert.True(document.Shapes.Count == 2);
                Assert.True(document.ParagraphsShapes.Count == 2);
                Assert.True(document.Sections[0].Shapes.Count == 2);
                Assert.True(document.Sections[0].ParagraphsShapes.Count == 2);

                rect.Remove();

                Assert.True(document.Shapes.Count == 1);
                Assert.Equal(Color.Yellow.ToHexColor(), document.Shapes[0].FillColorHex);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Shapes.Count == 1);
                Assert.Equal(Color.Yellow.ToHexColor(), document.Shapes[0].FillColorHex);
            }
        }
    }
}
