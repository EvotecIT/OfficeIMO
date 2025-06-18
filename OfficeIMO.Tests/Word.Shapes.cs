using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithShapes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with shape");
                var shape = paragraph.AddShape(100, 50, "#FF0000");

                Assert.True(document.Paragraphs.Count == 1);
                Assert.NotNull(paragraph.Shape);
                Assert.Equal(100d, shape.Width, 1);
                Assert.Equal(50d, shape.Height, 1);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithShapes.docx"))) {
                Assert.True(document.Paragraphs[0].Shape != null);
                Assert.Equal(100d, document.Paragraphs[0].Shape.Width, 1);
                Assert.Equal(50d, document.Paragraphs[0].Shape.Height, 1);
            }
        }
    }
}
