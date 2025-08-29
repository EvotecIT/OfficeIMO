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
                var loadedShape = document.Paragraphs[0].Shape;
                Assert.NotNull(loadedShape);
                Assert.Equal("Rectangle", loadedShape!.Title);
                Assert.Equal("My rectangle", loadedShape.Description);
                Assert.False(loadedShape.Hidden!.Value);
                Assert.True(loadedShape.Stroked!.Value);
                Assert.Equal(Color.Blue.ToHexColor(), loadedShape.StrokeColorHex);
                Assert.Equal(2d, loadedShape.StrokeWeight!.Value, 1);
                Assert.Equal(120d, loadedShape.Width, 1);
                Assert.Equal(60d, loadedShape.Height, 1);
                Assert.Equal(10d, loadedShape.Left!.Value, 1);
                Assert.Equal(20d, loadedShape.Top!.Value, 1);
                Assert.Equal(45d, loadedShape.Rotation!.Value, 1);
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
                var loadedShape = document.Paragraphs[0].Shape;
                Assert.NotNull(loadedShape);
                Assert.Equal(Color.Lime.ToHexColor(), loadedShape!.FillColorHex);
                Assert.Equal(Color.Black.ToHexColor(), loadedShape.StrokeColorHex);
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
                var loadedShape = document.Paragraphs[0].Shape;
                Assert.NotNull(loadedShape);
                Assert.Equal(Color.Aqua.ToHexColor(), loadedShape!.FillColorHex);
                Assert.Equal(Color.Red.ToHexColor(), loadedShape.StrokeColorHex);
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

        [Fact]
        public void Test_ShapesInSectionAndHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionHeaderShapes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var section = document.Sections[0];
                section.AddShape(ShapeType.Rectangle, 40, 20, Color.Red, Color.Black);
                section.AddShape(ShapeType.RoundedRectangle, 30, 15, Color.Yellow, Color.Black, 1, arcSize: 0.3);
                section.AddShapeDrawing(ShapeType.Ellipse, 20, 20);

                section.AddHeadersAndFooters();
                section.Header.Default.AddShape(ShapeType.Rectangle, 30, 15, Color.Blue, Color.Black);
                section.Header.Default.AddShape(ShapeType.RoundedRectangle, 25, 15, Color.Green, Color.Black, 1, arcSize: 0.3);
                section.Header.Default.AddShapeDrawing(ShapeType.Ellipse, 20, 20);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var section = document.Sections[0];
                Assert.Equal(3, document.Shapes.Count);
                Assert.Equal(3, section.Shapes.Count);
                Assert.NotNull(section.Header);
                Assert.NotNull(section.Header!.Default);
                var headerDefault = section.Header.Default!;
                Assert.True(headerDefault.Paragraphs[0].IsShape);
                Assert.True(headerDefault.Paragraphs[1].IsShape);
                Assert.True(headerDefault.Paragraphs[2].IsShape);
            }
        }

        [Fact]
        public void Test_AddRoundedRectangleShape() {
            string filePath = Path.Combine(_directoryWithFiles, "RoundedRectangleShape.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var shape = document.AddShape(ShapeType.RoundedRectangle, 60, 30, Color.Lime, Color.Black, 1, arcSize: 0.3);
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.NotNull(shape.ArcSize);
                Assert.InRange(shape.ArcSize!.Value, 0.29, 0.31);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].IsShape);
                var loadedShape = document.Paragraphs[0].Shape;
                Assert.NotNull(loadedShape);
                Assert.InRange(loadedShape!.ArcSize!.Value, 0.29, 0.31);
            }
        }

        [Fact]
        public void Test_WordShape_InternalProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "ShapeInternalProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph();
                var rectangle = paragraph.AddShape(100, 50);
                var line = WordShape.AddLine(paragraph, 0, 0, 10, 10);

                Assert.NotNull(rectangle.Run);
                Assert.Null(rectangle.Line);

                Assert.NotNull(line.Run);
                Assert.NotNull(line.Line);
            }
        }
    }
}
