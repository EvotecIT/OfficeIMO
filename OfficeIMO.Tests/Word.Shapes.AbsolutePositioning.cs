using System;
using System.Linq;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void AddShapeDrawing_WithValidCoordinates_CreatesAnchoredShape() {
            string filePath = Path.Combine(_directoryWithFiles, "AnchoredShape_Valid.docx");
            using (var document = WordDocument.Create(filePath)) {
                var p = document.AddParagraph("Anchor here");
                var shape = WordShape.AddDrawingShapeAnchored(p, ShapeType.Rectangle, 80, 40, leftPt: 36, topPt: 72);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                // Verify there's a drawing anchor in the first paragraph
                var p = document.Paragraphs.First();
                Assert.True(p.IsSmartArt == false); // sanity check
                Assert.NotNull(p._run);
                var hasAnchor = p._run!.ChildElements.OfType<DocumentFormat.OpenXml.Wordprocessing.Drawing>()
                    .Any(d => d.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>().Any());
                Assert.True(hasAnchor);
            }
        }

        [Fact]
        public void AddShapeDrawing_WithNegativeCoordinates_ThrowsException() {
            using var document = WordDocument.Create();
            var p = document.AddParagraph();
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, ShapeType.Rectangle, 10, 10, -1, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, ShapeType.Rectangle, 10, 10, 0, -5));
        }

        [Fact]
        public void AddShapeDrawing_WithExtremeValues_ThrowsException() {
            using var document = WordDocument.Create();
            var p = document.AddParagraph();
            // Use a very large value to trigger EMU overflow guard
            double huge = double.MaxValue / 2; // will overflow when multiplied
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, ShapeType.Rectangle, 10, 10, huge, 0));
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, ShapeType.Rectangle, 10, 10, 0, huge));
        }

        [Fact]
        public void AddShapeDrawing_MultipleShapes_DoesNotInterfere() {
            string filePath = Path.Combine(_directoryWithFiles, "AnchoredShape_Multiple.docx");
            using (var document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("A");
                var p2 = document.AddParagraph("B");
                WordShape.AddDrawingShapeAnchored(p1, ShapeType.Rectangle, 40, 40, 10, 10);
                WordShape.AddDrawingShapeAnchored(p2, ShapeType.Ellipse, 30, 30, 20, 20);
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var anchors = document.Paragraphs.SelectMany(q => q._run?.ChildElements
                    .OfType<DocumentFormat.OpenXml.Wordprocessing.Drawing>() ?? Enumerable.Empty<DocumentFormat.OpenXml.Wordprocessing.Drawing>())
                    .SelectMany(d => d.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>())
                    .Count();
                Assert.Equal(2, anchors);
            }
        }
    }
}

