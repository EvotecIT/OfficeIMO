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
                var shape = WordShape.AddDrawingShapeAnchored(p, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(40), left: OfficeIMO.OfficeLength.Points(36), top: OfficeIMO.OfficeLength.Points(72));
                document.Save();
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
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(-1), OfficeIMO.OfficeLength.Points(0)));
            Assert.Throws<ArgumentOutOfRangeException>(() => WordShape.AddDrawingShapeAnchored(p, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(-5)));
        }

        [Fact]
        public void AddShapeDrawing_WithExtremeValues_ThrowsException() {
            using var document = WordDocument.Create();
            var p = document.AddParagraph();
            // Length construction rejects values beyond the shared whole-EMU range.
            double huge = double.MaxValue / 2; // will overflow when multiplied
            Assert.Throws<OverflowException>(() => WordShape.AddDrawingShapeAnchored(p, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(huge), OfficeIMO.OfficeLength.Points(0)));
            Assert.Throws<OverflowException>(() => WordShape.AddDrawingShapeAnchored(p, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(huge)));
        }

        [Fact]
        public void AddShapeDrawing_MultipleShapes_DoesNotInterfere() {
            string filePath = Path.Combine(_directoryWithFiles, "AnchoredShape_Multiple.docx");
            using (var document = WordDocument.Create(filePath)) {
                var p1 = document.AddParagraph("A");
                var p2 = document.AddParagraph("B");
                WordShape.AddDrawingShapeAnchored(p1, WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(40), OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10));
                WordShape.AddDrawingShapeAnchored(p2, WordShapeType.Ellipse, OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(30), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(20));
                document.Save();
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

