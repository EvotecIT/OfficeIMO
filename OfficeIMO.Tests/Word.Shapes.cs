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

                WordShape.AddEllipse(paragraph, 80, 40, Color.Lime);
                WordShape.AddPolygon(paragraph, "0,0 50,0 50,50 0,50", Color.White, Color.Blue);
                WordShape.AddLine(paragraph, 0, 60, 100, 60, Color.Red, 2);

                Assert.True(document.Paragraphs.Count == 1);
                Assert.NotNull(paragraph.Shape);
                Assert.Equal(100d, shape.Width, 1);
                Assert.Equal(50d, shape.Height, 1);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithShapes.docx"))) {
                Assert.True(document.Paragraphs[0].IsShape);
                Assert.Equal(100d, document.Paragraphs[0].Shape.Width, 1);
                Assert.Equal(50d, document.Paragraphs[0].Shape.Height, 1);
            }
        }
    }
}
