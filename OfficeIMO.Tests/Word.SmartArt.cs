using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        public static IEnumerable<object[]> SmartArtTypesWithLayouts() {
            yield return new object[] { SmartArtType.BasicProcess, "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            yield return new object[] { SmartArtType.Hierarchy, "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1" };
            yield return new object[] { SmartArtType.Cycle, "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" };
            yield return new object[] { SmartArtType.PictureOrgChart, "urn:microsoft.com/office/officeart/2005/8/layout/pictureorgchart" };
            yield return new object[] { SmartArtType.ContinuousBlockProcess, "urn:microsoft.com/office/officeart/2005/8/layout/process6" };
            yield return new object[] { SmartArtType.CustomSmartArt1, "urn:microsoft.com/office/officeart/2005/8/layout/default" };
            yield return new object[] { SmartArtType.CustomSmartArt2, "urn:microsoft.com/office/officeart/2005/8/layout/cycle2" };
        }

        [Fact]
        public void Test_AddSmartArt() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtDocument.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddSmartArt(SmartArtType.BasicProcess);
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                Assert.Single(mainPart.DiagramDataParts);
                Assert.Single(mainPart.DiagramLayoutDefinitionParts);
                Assert.Single(mainPart.DiagramStyleParts);
                Assert.Single(mainPart.DiagramColorsParts);
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                Assert.Single(mainPart.DiagramDataParts);
                Assert.Single(mainPart.DiagramLayoutDefinitionParts);
                Assert.Single(mainPart.DiagramStyleParts);
                Assert.Single(mainPart.DiagramColorsParts);
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
            }
        }

        [Fact]
        public void Test_SmartArt_Retrieval_After_Load() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtRetrieve.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddSmartArt(SmartArtType.Hierarchy);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.SmartArts);
                Assert.Single(document.Sections[0].SmartArts);
                Assert.Single(document.ParagraphsSmartArts);
            }
        }

        [Fact]
        public void Test_SmartArt_Text_Edit_Persists() {
            string filePath = Path.Combine(_directoryWithFiles, "SmartArtEdit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var sa = document.AddSmartArt(SmartArtType.BasicProcess);
                // Debug: ensure the SmartArt data part contains at least 1 editable node
                var main = document._wordprocessingDocument.MainDocumentPart!;
                var rel = sa._drawing.Descendants<DocumentFormat.OpenXml.Drawing.Diagrams.RelationshipIds>().First();
                var dataPart = (DocumentFormat.OpenXml.Packaging.DiagramDataPart)main.GetPartById(rel.DataPart!);
                using (var s = dataPart.GetStream(FileMode.Open, FileAccess.Read)) {
                    var xdoc = System.Xml.Linq.XDocument.Load(s);
                    var dgm = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/diagram";
                    var a = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
                    var nodePts = xdoc.Descendants(dgm + "pt").Where(p => p.Attribute("type") == null && (p.Element(dgm + "t") != null || p.Element(dgm + "txBody") != null)).ToList();
                    var paras = nodePts.Select(p => (p.Element(dgm + "t") ?? p.Element(dgm + "txBody"))?.Element(a + "p")).Where(p => p != null).ToList();
                    if (paras.Count < 1) {
                        var debugDir = Path.Combine(_directoryWithFiles, "_debug");
                        Directory.CreateDirectory(debugDir);
                        var outPath = Path.Combine(debugDir, "diagram.xml");
                        File.WriteAllText(outPath, xdoc.ToString());
                        Assert.True(paras.Count >= 1, $"Expected at least 1 node paragraph, got {paras.Count}. Data written to: {outPath}");
                    }
                }
                sa.SetNodeText(0, "Hello");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var sa = document.SmartArts.Single();
                Assert.Equal("Hello", sa.GetNodeText(0));
            }
        }

        [Theory]
        [MemberData(nameof(SmartArtTypesWithLayouts))]
        public void Test_SmartArt_Relationships_For_Each_Type(SmartArtType type, string expectedLayoutUid) {
            string filePath = Path.Combine(_directoryWithFiles, $"SmartArt_{type}.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var smartArt = document.AddSmartArt(type);
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                var rel = smartArt._drawing.Descendants<RelationshipIds>().First();

                Assert.False(string.IsNullOrEmpty(rel.LayoutPart));
                Assert.False(string.IsNullOrEmpty(rel.ColorPart));
                Assert.False(string.IsNullOrEmpty(rel.StylePart));
                Assert.False(string.IsNullOrEmpty(rel.DataPart));

                var layoutPart = (DiagramLayoutDefinitionPart)mainPart.GetPartById(rel.LayoutPart!);
                var colorsPart = (DiagramColorsPart)mainPart.GetPartById(rel.ColorPart!);
                var stylePart = (DiagramStylePart)mainPart.GetPartById(rel.StylePart!);
                var dataPart = (DiagramDataPart)mainPart.GetPartById(rel.DataPart!);

                Assert.Equal(expectedLayoutUid, layoutPart.LayoutDefinition?.UniqueId?.Value);
                Assert.NotNull(colorsPart.ColorsDefinition);
                Assert.NotNull(stylePart.StyleDefinition);
                using (var dataStream = dataPart.GetStream()) {
                    Assert.True(dataStream.Length > 0);
                }

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var smartArt = document.SmartArts.Single();
                var mainPart = document._wordprocessingDocument.MainDocumentPart!;
                var rel = smartArt._drawing.Descendants<RelationshipIds>().First();
                var layoutPart = (DiagramLayoutDefinitionPart)mainPart.GetPartById(rel.LayoutPart!);

                Assert.Equal(expectedLayoutUid, layoutPart.LayoutDefinition?.UniqueId?.Value);
            }
        }
    }
}
