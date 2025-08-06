using System.IO;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DrawingVsVmlCounts() {
            string assets = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../..", "Assets"));
            string img = Path.Combine(assets, "OfficeIMO.png");

            string drawingFile = Path.Combine(_directoryWithFiles, "DrawingCounts.docx");
            using (WordDocument doc = WordDocument.Create(drawingFile)) {
                doc.AddParagraph().AddImage(img);
                doc.AddShapeDrawing(ShapeType.Ellipse, 40, 40);
                doc.AddShapeDrawing(ShapeType.Rectangle, 60, 30);
                doc.AddTextBox("Text");
                doc.Save(false);
            }
            using (WordDocument doc = WordDocument.Load(drawingFile)) {
                Assert.Single(doc.Images);
                Assert.Equal(2, doc.Shapes.Count);
                Assert.Single(doc.TextBoxes);
            }

            string vmlFile = Path.Combine(_directoryWithFiles, "VmlCounts.docx");
            using (WordDocument doc = WordDocument.Create(vmlFile)) {
                doc.AddImageVml(img);
                doc.AddShape(ShapeType.Ellipse, 40, 40, Color.Red, Color.Blue);
                doc.AddShape(ShapeType.Rectangle, 60, 30, Color.Green, Color.Black);
                doc.AddTextBoxVml("Text");
                doc.Save(false);
            }
            using (WordDocument doc = WordDocument.Load(vmlFile)) {
                Assert.Single(doc.Images);
                Assert.Equal(2, doc.Shapes.Count);
                Assert.Single(doc.TextBoxes);
            }
        }
    }
}

