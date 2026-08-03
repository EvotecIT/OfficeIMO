using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ShapeGroup_PersistsBoundedChildrenColorsAndInlineLayout() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "ShapeGroup.Inline.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph();
                WordShapeGroup group = paragraph.AddShapeGroup(new[] {
                    new WordShapeGroupItem(ShapeType.Rectangle, 0, 0, 72, 36) {
                        FillColorHex = "#336699",
                        StrokeColorHex = "112233"
                    },
                    new WordShapeGroupItem(ShapeType.Ellipse, 90, 18, 54, 54) {
                        FillColorHex = "F0A000"
                    }
                });

                Assert.Equal(2, group.ChildCount);
                Assert.True(paragraph.IsShapeGroup);
                Assert.False(paragraph.IsShape);
                Assert.True(group.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
                Assert.Equal(WordDrawingPlacementKind.Inline, layout.Placement);
                Assert.True(layout.IsGroup);
                Assert.Equal(144D, layout.WidthPoints, 6);
                Assert.Equal(72D, layout.HeightPoints, 6);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordParagraph importedParagraph = Assert.Single(reloaded.Paragraphs);
            WordShapeGroup imported = Assert.IsType<WordShapeGroup>(importedParagraph.ShapeGroup);
            Assert.Equal(2, imported.ChildCount);
            Assert.False(importedParagraph.IsShape);
            Wpg.WordprocessingGroup xmlGroup = importedParagraph._run!
                .Descendants<Wpg.WordprocessingGroup>().Single();
            Wps.WordprocessingShape first = xmlGroup.Descendants<Wps.WordprocessingShape>().First();
            Assert.Equal("336699", first.Descendants<RgbColorModelHex>().First().Val!.Value);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(reloaded._wordprocessingDocument));
        }

        [Fact]
        public void ShapeGroup_PersistsAnchoredPlacementWithoutFlatteningChildren() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "ShapeGroup.Anchored.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordShapeGroup group = document.AddParagraph().AddShapeGroup(new[] {
                    new WordShapeGroupItem(ShapeType.Chevron, 0, 0, 80, 40),
                    new WordShapeGroupItem(ShapeType.Chevron, 72, 0, 80, 40),
                    new WordShapeGroupItem(ShapeType.Chevron, 144, 0, 80, 40)
                }, 24, 48);
                Assert.True(group.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
                Assert.Equal(WordDrawingPlacementKind.Anchored, layout.Placement);
                Assert.Equal(24D, layout.HorizontalOffsetPoints!.Value, 6);
                Assert.Equal(48D, layout.VerticalOffsetPoints!.Value, 6);
                Assert.Equal(WordDrawingWrapKind.Square, layout.Wrap);
                Assert.True(layout.IsGroup);
                document.Save();
            }

            using WordDocument reloaded = WordDocument.Load(filePath);
            WordShapeGroup imported = Assert.IsType<WordShapeGroup>(Assert.Single(reloaded.Paragraphs).ShapeGroup);
            Assert.Equal(3, imported.ChildCount);
            Assert.True(imported.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot persisted));
            Assert.Equal(WordDrawingPlacementKind.Anchored, persisted.Placement);
            Assert.Equal(224D, persisted.WidthPoints, 6);
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2010)
                .Validate(reloaded._wordprocessingDocument));
        }
    }
}
