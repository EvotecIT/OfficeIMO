using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void ShapeGroup_PersistsBoundedChildrenColorsAndInlineLayout() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "ShapeGroup.Inline.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordParagraph paragraph = document.AddParagraph();
                WordShapeGroup group = paragraph.AddShapeGroup(new[] {
                    new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 72, 36) {
                        FillColorHex = "#336699",
                        StrokeColorHex = "112233"
                    },
                    new WordShapeGroupItem(WordShapeType.Ellipse, 90, 18, 54, 54) {
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
                    new WordShapeGroupItem(WordShapeType.Chevron, 0, 0, 80, 40),
                    new WordShapeGroupItem(WordShapeType.Chevron, 72, 0, 80, 40),
                    new WordShapeGroupItem(WordShapeType.Chevron, 144, 0, 80, 40)
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

        [Fact]
        public void DrawingShapes_AllocateIdsAboveExistingPackageIdsAfterReload() {
            string filePath = System.IO.Path.Combine(_directoryWithFiles, "ShapeGroup.DocumentScopedIds.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordShape existing = WordShape.AddDrawingShape(document.AddParagraph(), WordShapeType.Rectangle, 40, 20);
                existing._drawing!.Inline!.DocProperties!.Id = 100U;
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph().AddShapeGroup(new[] {
                    new WordShapeGroupItem(WordShapeType.Rectangle, 0, 0, 20, 20),
                    new WordShapeGroupItem(WordShapeType.Ellipse, 30, 0, 20, 20),
                });
                WordShape.AddDrawingShape(document.AddParagraph(), WordShapeType.Diamond, 30, 30);
                WordShape.AddDrawingShapeAnchored(document.AddParagraph(), WordShapeType.Chevron, 30, 20, 10, 10);
                WordChart chart = document.AddChart("ID allocation");
                chart.AddCategories(new List<string> { "A" });
                chart.AddBar("Series", new List<int> { 1 }, OfficeIMO.Drawing.OfficeColor.Blue);
                document.AddSmartArt(WordSmartArtType.BasicProcess);
                document.AddParagraph().AddImage(
                    System.IO.Path.Combine(_directoryWithImages, "EvotecLogo.png"),
                    20,
                    20);
                document.AddParagraph().AddImage(
                    System.IO.Path.Combine(_directoryWithImages, "EvotecLogo.png"),
                    20,
                    20);
                document.AddTextBox("Allocated text box");
                document.AddHeadersAndFooters();
                RequireSectionHeader(document, 0, DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default)
                    .AddPageNumber(WordPageNumberStyle.VerticalOutline2);

                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                IEnumerable<OpenXmlElement> roots = new[] { (OpenXmlElement)mainPart.Document! }
                    .Concat(mainPart.HeaderParts.Select(part => (OpenXmlElement)part.Header!))
                    .Concat(mainPart.FooterParts.Select(part => (OpenXmlElement)part.Footer!));
                uint[] ids = roots.SelectMany(root => root.Descendants<DW.DocProperties>()).Select(properties => properties.Id!.Value)
                    .Concat(roots.SelectMany(root => root.Descendants<PIC.NonVisualDrawingProperties>()).Select(properties => properties.Id!.Value))
                    .Concat(roots.SelectMany(root => root.Descendants<Wpg.NonVisualDrawingProperties>()).Select(properties => properties.Id!.Value))
                    .Concat(roots.SelectMany(root => root.Descendants<Wps.NonVisualDrawingProperties>()).Select(properties => properties.Id!.Value))
                    .ToArray();
                Assert.Equal(ids.Length, ids.Distinct().Count());
                Assert.All(ids.Where(id => id != 100U), id => Assert.True(id > 100U));
                var validationErrors = new OpenXmlValidator(FileFormatVersions.Office2010)
                    .Validate(document._wordprocessingDocument)
                    .Where(error => error.Id != "Sch_AttributeValueDataTypeDetailed" ||
                                    error.Description?.Contains("attribute 'title' has invalid value ''", StringComparison.Ordinal) != true)
                    .ToList();
                Assert.Empty(validationErrors);
            }
        }
    }
}
