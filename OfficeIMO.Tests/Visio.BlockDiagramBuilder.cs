using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioBlockDiagramBuilderTests {
        [Fact]
        public void BlockDiagramBuilderCreatesRegionsBlocksAndFlows() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .BlockDiagram("System", diagram => diagram
                    .Region("compute", "Compute", 1, 1, 2, 2)
                    .Block("input", "Input Device", 0, 1)
                    .EmphasisBlock("memory", "Memory Unit", 1, 1)
                    .Block("storage", "Secondary\nStorage", 1, 0, VisioBlockShapeKind.Data)
                    .Block("output", "Output Device", 3, 1)
                    .DataFlow("input", "memory")
                    .DataFlow("memory", "output")
                    .ControlFlow("storage", "output", "Control"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("System", page.Name);
            Assert.Equal(5, page.Shapes.Count);
            Assert.Equal(new[] { "compute", "input", "memory", "storage", "output" }, page.Shapes.Select(shape => shape.Id).ToArray());
            Assert.Equal("Rectangle", page.Shapes[0].NameU);
            Assert.Equal("Process", page.Shapes[1].NameU);
            Assert.Equal("Process", page.Shapes[2].NameU);
            Assert.Equal("Data", page.Shapes[3].NameU);
            Assert.Equal("Process", page.Shapes[4].NameU);
            Assert.Equal(3, page.Connectors.Count);
            Assert.Equal(2, page.Connectors.Count(connector => connector.LinePattern == 1));
            Assert.Single(page.Connectors, connector => connector.LinePattern == 2);
            Assert.Contains(page.Connectors, connector => connector.Label == "Control");
            Assert.All(page.Connectors, connector => Assert.Equal(EndArrow.Triangle, connector.EndArrow));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(5, loaded.Pages[0].Shapes.Count);
            Assert.Equal(3, loaded.Pages[0].Connectors.Count);
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "Control");
        }

        [Fact]
        public void BlockDiagramBuilderRejectsDuplicateRegionAndBlockIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Region("same", "Region", 0, 0, 1, 1)
                    .Block("same", "Block", 0, 0)));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void BlockDiagramBuilderRejectsUnknownFlowEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("known", "Known", 0, 0)
                    .DataFlow("known", "missing")));

            Assert.Contains("Unknown block id", exception.Message);
        }

        [Fact]
        public void BlockDiagramBuilderCanAddTitleAndLegendWithoutOverlappingGrid() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .BlockDiagram("Block Diagram", diagram => diagram
                    .Title()
                    .Legend()
                    .Block("input", "Input", 0, 0)
                    .Block("memory", "Memory", 1, 0)
                    .Block("output", "Output", 2, 0)
                    .DataFlow("input", "memory")
                    .ControlFlow("memory", "output", "Control"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Contains(page.Shapes, shape => shape.Text == "Block Diagram" && shape.NameU == "Text Box");
            Assert.Contains(page.Shapes, shape => shape.Text == "Data Flow");
            Assert.Contains(page.Shapes, shape => shape.Text == "Control Flow");
            Assert.True(page.Shapes.Single(shape => shape.Id == "input").PinY < page.Height - 1.2D);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }
    }
}
