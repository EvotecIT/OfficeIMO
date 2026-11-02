using System.IO;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_EllipseAndPolygonShapes() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithEllipsePolygon.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with shapes");
                var ellipse = WordShape.AddEllipse(paragraph, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(20), Color.Lime);
                WordShape.AddPolygon(paragraph, "0,0 20,0 20,20 0,20", Color.Yellow, Color.Blue);

                Assert.NotNull(ellipse);
                Assert.Equal(40d, ellipse.Width.ToPoints(), 1);
                Assert.Equal(20d, ellipse.Height.ToPoints(), 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].IsShape);
                var shape = document.Paragraphs[0].Shape;
                Assert.NotNull(shape);
                Assert.Equal(40d, shape!.Width.ToPoints(), 1);
                Assert.Equal(20d, shape.Height.ToPoints(), 1);
            }
        }
        [Fact]
        public void Test_LineShapeFactory() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithLineFactory.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Line via factory");
                WordShape.AddLine(paragraph, OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(50), OfficeIMO.OfficeLength.Points(0), Color.Red, OfficeIMO.OfficeLength.Points(1));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.Paragraphs[0].IsShape);
            }
        }
    }
}
