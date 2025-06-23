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
    }
}
