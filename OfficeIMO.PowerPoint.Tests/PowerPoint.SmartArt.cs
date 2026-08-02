using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSmartArtTests {
        [Fact]
        public void CanAddSmartArtAndEditNodeText() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointSmartArt smartArt = slide.AddSmartArt();

                    Assert.Equal(PowerPointShapeContentType.SmartArt, smartArt.ShapeContentType);
                    Assert.Equal(1, smartArt.NodeCount);
                    smartArt.SetNodeText(0, "OfficeIMO-native process");
                    Assert.Equal("OfficeIMO-native process", smartArt.GetNodeText(0));
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Assert.Single(slidePart.DiagramDataParts);
                    Assert.Single(slidePart.DiagramLayoutDefinitionParts);
                    Assert.Single(slidePart.DiagramStyleParts);
                    Assert.Single(slidePart.DiagramColorsParts);

                    GraphicFrame frame = slidePart.Slide.Descendants<GraphicFrame>().Single();
                    RelationshipIds relationships = frame.Graphic!.GraphicData!.GetFirstChild<RelationshipIds>()!;
                    Assert.False(string.IsNullOrWhiteSpace(relationships.LayoutPart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.StylePart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.ColorPart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.DataPart));

                    DiagramLayoutDefinitionPart layoutPart =
                        (DiagramLayoutDefinitionPart)slidePart.GetPartById(relationships.LayoutPart!);
                    Assert.Equal("urn:microsoft.com/office/officeart/2005/8/layout/process1",
                        layoutPart.LayoutDefinition!.UniqueId!.Value);
                }

                using (PowerPointPresentation reloaded = PowerPointPresentation.Load(filePath)) {
                    PowerPointSmartArt smartArt = Assert.IsType<PowerPointSmartArt>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal("OfficeIMO-native process", smartArt.GetNodeText(0));
                    Assert.Single(reloaded.Slides[0].SmartArts);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(long.MaxValue, 0L)]
        [InlineData(long.MinValue, 0L)]
        [InlineData(0L, long.MaxValue)]
        [InlineData(0L, long.MinValue)]
        public void AddSmartArtRejectsOffsetsOutsideDrawingRangeBeforeMutation(
            long left, long top) {
            string filePath = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();

                    Assert.Throws<ArgumentOutOfRangeException>(() =>
                        slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                            new[] { "Plan", "Build" }, left, top,
                            5486400L, 3200400L));

                    Assert.Empty(slide.Shapes);
                    presentation.Save();
                }

                using PresentationDocument document =
                    PresentationDocument.Open(filePath, false);
                SlidePart slidePart = document.PresentationPart!
                    .SlideParts.Single();
                Assert.Empty(slidePart.DiagramDataParts);
                Assert.Empty(slidePart.DiagramLayoutDefinitionParts);
                Assert.Empty(slidePart.DiagramStyleParts);
                Assert.Empty(slidePart.DiagramColorsParts);
                Assert.Empty(slidePart.Slide.Descendants<GraphicFrame>());
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ImportedSmartArtPreservesEveryNodeParagraphInSemanticExports() {
            string filePath = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(filePath)) {
                    PowerPointSmartArt authored = presentation.AddSlide()
                        .AddSmartArt();
                    authored.SetNodeText(0, "First paragraph");
                    presentation.Save();
                }

                using (PresentationDocument document =
                       PresentationDocument.Open(filePath, true)) {
                    DiagramDataPart dataPart = document.PresentationPart!
                        .SlideParts.Single().DiagramDataParts.Single();
                    XDocument data;
                    using (Stream input = dataPart.GetStream(
                               FileMode.Open, FileAccess.Read)) {
                        data = XDocument.Load(input);
                    }
                    XNamespace dgm =
                        "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                    XNamespace a =
                        "http://schemas.openxmlformats.org/drawingml/2006/main";
                    XElement textBody = data.Descendants(dgm + "pt")
                        .Where(point => point.Attribute("type") == null)
                        .Select(point => point.Element(dgm + "t")
                            ?? point.Element(dgm + "txBody"))
                        .First(body => body != null)!;
                    XElement firstParagraph = textBody.Elements(a + "p")
                        .Single();
                    XElement firstRun = firstParagraph.Elements(a + "r")
                        .Single();
                    firstRun.AddAfterSelf(new XElement(a + "br"),
                        new XElement(a + "r",
                            new XElement(a + "t", "After break")));
                    textBody.Add(new XElement(a + "p",
                        new XElement(a + "r",
                            new XElement(a + "t", "Second paragraph")),
                        new XElement(a + "endParaRPr",
                            new XAttribute("lang", "en-US"))));
                    using Stream output = dataPart.GetStream(
                        FileMode.Create, FileAccess.Write);
                    data.Save(output);
                }

                using PowerPointPresentation imported =
                    PowerPointPresentation.Load(filePath);
                PowerPointSmartArt smartArt = Assert.Single(
                    imported.Slides[0].SmartArts);
                Assert.Equal(1, smartArt.NodeCount);
                Assert.Equal("First paragraph\nAfter break\nSecond paragraph",
                    smartArt.GetNodeText(0));
                Assert.Equal("First paragraph\nAfter break\nSecond paragraph",
                    Assert.Single(smartArt.GetNodeTexts()));
                Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot snapshot));
                Assert.Equal("First paragraph\nAfter break\nSecond paragraph",
                    Assert.Single(snapshot.Nodes));

                OfficeImageExportResult svg = imported.Slides[0].ExportImage(
                    OfficeImageExportFormat.Svg);
                string svgText = Encoding.UTF8.GetString(svg.Bytes);
                Assert.Contains("First paragraph", svgText,
                    StringComparison.Ordinal);
                Assert.Contains("After break", svgText,
                    StringComparison.Ordinal);
                Assert.Contains("Second paragraph", svgText,
                    StringComparison.Ordinal);

                smartArt.SetNodeText(0, "Replacement");
                Assert.Equal("Replacement", smartArt.GetNodeText(0));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ImportedSmartArtUsesConnectionOrderForSemanticNodes() {
            string filePath = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddSmartArt(
                        PowerPointSmartArtType.BasicProcess,
                        new[] { "First", "Second", "Third" });
                    presentation.Save();
                }

                using (PresentationDocument document =
                       PresentationDocument.Open(filePath, true)) {
                    DiagramDataPart dataPart = document.PresentationPart!
                        .SlideParts.Single().DiagramDataParts.Single();
                    XDocument data;
                    using (Stream input = dataPart.GetStream(
                               FileMode.Open, FileAccess.Read)) {
                        data = XDocument.Load(input);
                    }
                    XNamespace dgm =
                        "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                    XElement pointList = data.Descendants(dgm + "ptLst")
                        .Single();
                    XElement[] nodes = pointList.Elements(dgm + "pt")
                        .Where(point => point.Attribute("type") == null)
                        .ToArray();
                    for (int index = nodes.Length - 1; index >= 0; index--) {
                        XElement node = nodes[index];
                        node.Remove();
                        pointList.Add(node);
                    }
                    using Stream output = dataPart.GetStream(
                        FileMode.Create, FileAccess.Write);
                    data.Save(output);
                }

                using PowerPointPresentation imported =
                    PowerPointPresentation.Load(filePath);
                PowerPointSmartArt smartArt = Assert.Single(
                    imported.Slides[0].SmartArts);
                Assert.True(smartArt.TryGetOfficeDiagramSnapshot(
                    out OfficeDiagramSnapshot snapshot));
                Assert.Equal(new[] { "First", "Second", "Third" },
                    snapshot.Nodes);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ImportedSmartArtRejectsTopologyThatSemanticRendererCannotRepresent() {
            string filePath = Path.Combine(Path.GetTempPath(),
                Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddSmartArt(
                        PowerPointSmartArtType.BasicHierarchy,
                        new[] { "Root", "Child", "Grandchild" });
                    presentation.Save();
                }

                using (PresentationDocument document =
                       PresentationDocument.Open(filePath, true)) {
                    DiagramDataPart dataPart = document.PresentationPart!
                        .SlideParts.Single().DiagramDataParts.Single();
                    XDocument data;
                    using (Stream input = dataPart.GetStream(
                               FileMode.Open, FileAccess.Read)) {
                        data = XDocument.Load(input);
                    }
                    XNamespace dgm =
                        "http://schemas.openxmlformats.org/drawingml/2006/diagram";
                    XElement[] nodePoints = data.Descendants(dgm + "pt")
                        .Where(point => point.Attribute("type") == null)
                        .ToArray();
                    string childId = (string)nodePoints[1]
                        .Attribute("modelId")!;
                    string grandchildId = (string)nodePoints[2]
                        .Attribute("modelId")!;
                    XElement grandchildConnection = data
                        .Descendants(dgm + "cxn")
                        .Single(connection => string.Equals(
                            (string?)connection.Attribute("destId"),
                            grandchildId, StringComparison.Ordinal));
                    grandchildConnection.SetAttributeValue("srcId", childId);
                    using Stream output = dataPart.GetStream(
                        FileMode.Create, FileAccess.Write);
                    data.Save(output);
                }

                using PowerPointPresentation imported =
                    PowerPointPresentation.Load(filePath);
                PowerPointSmartArt smartArt = Assert.Single(
                    imported.Slides[0].SmartArts);
                Assert.False(smartArt.TryGetOfficeDiagramSnapshot(out _));

                OfficeImageExportResult svg = imported.Slides[0].ExportImage(
                    OfficeImageExportFormat.Svg);
                Assert.Contains(svg.Diagnostics, diagnostic =>
                    diagnostic.Message.Contains("semantic node data",
                        StringComparison.OrdinalIgnoreCase));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
