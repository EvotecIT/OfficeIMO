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
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            double highestGridTop = page.Shapes
                .Where(shape => shape.Id != "title" && !string.Equals(shape.Text, "Data Flow", StringComparison.Ordinal) && !string.Equals(shape.Text, "Control Flow", StringComparison.Ordinal))
                .Max(shape => shape.PinY + shape.Height / 2D);
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Block Diagram", title.Text);
            Assert.True(title.PinY - title.Height / 2D > highestGridTop);
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

        [Fact]
        public void BlockDiagramBuilderNormalizesFlowEndpointIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .BlockDiagram("Trimmed", diagram => diagram
                    .Block("input", "Input", 0, 0)
                    .Block("output", "Output", 1, 0)
                    .DataFlow(" input ", " output ", "data"));

            VisioPage page = Assert.Single(document.Pages);
            VisioConnector connector = Assert.Single(page.Connectors);
            Assert.Equal("input", connector.From.Id);
            Assert.Equal("output", connector.To.Id);
            Assert.Equal("data", connector.Label);
        }

        [Fact]
        public void BlockDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .BlockDiagram("Annotated Blocks", diagram => diagram
                    .Title()
                    .Region("compute", "Compute", 0, 0, 2, 1)
                    .Block("input", "Input", 0, 0)
                    .EmphasisBlock("processor", "Processor", 1, 0)
                    .DataFlow("input", "processor")
                    .Callout(" processor ", "processor-note", "High throughput path", 5.8, 5.9, options => {
                        options.Width = 2.4;
                        options.Height = 0.7;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "processor");
            Assert.Equal("processor-note", callout.Id);
            Assert.Equal("High throughput path", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.4, callout.Width);
            Assert.Equal(0.7, callout.Height);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, callout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void BlockDiagramBuilderCanAutoPlaceSemanticCalloutsBesideBlocks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .BlockDiagram("Auto Annotated Blocks", diagram => diagram
                    .Title()
                    .Region("compute", "Compute", 0, 0, 2, 1)
                    .Block("input", "Input", 0, 0)
                    .EmphasisBlock("processor", "Processor", 1, 0)
                    .DataFlow("input", "processor")
                    .Callout("processor", "processor-note", "High throughput path", VisioSide.Right, 0.4, options => {
                        options.Width = 2.4;
                        options.Height = 0.7;
                    })
                    .Callout("input", "Validate source", VisioSide.Left, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape processor = Assert.Single(page.Shapes, shape => shape.Id == "processor");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "processor-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "input-callout");

            Assert.True(explicitCallout.PinX > processor.PinX);
            Assert.Equal(processor.PinY, explicitCallout.PinY, 6);
            Assert.Equal(processor.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.4, explicitCallout.Width);
            Assert.True(generatedCallout.PinX < page.FindShapeById("input")!.PinX);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(processor, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void BlockDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .BlockDiagram("Generated", diagram => diagram
                    .Block("processor", "Processor", 0, 0)
                    .Callout("processor", "First note", 3.5, 4.5)
                    .Callout("processor", "Second note", 3.5, 3.6));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "processor-callout", "processor-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void BlockDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException blockFirst = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("title", "Input", 0, 0)
                    .Title()));
            ArgumentException titleFirst = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Title()
                    .Block("title", "Input", 0, 0)));
            ArgumentException regionCollision = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Region("title", "Region", 0, 0, 1, 1)
                    .Title()));

            Assert.Contains("already exists", blockFirst.Message);
            Assert.Contains("already exists", titleFirst.Message);
            Assert.Contains("already exists", regionCollision.Message);
        }

        [Fact]
        public void BlockDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("processor", "Processor", 0, 0)
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException blockCollision = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("processor", "Processor", 0, 0)
                    .Callout("processor", "processor", "Duplicate id", 4, 4)));
            ArgumentException regionCollision = Assert.Throws<ArgumentException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Region("compute", "Compute", 0, 0, 1, 1)
                    .Block("processor", "Processor", 0, 0)
                    .Callout("processor", "compute", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown block id", unknownTarget.Message);
            Assert.Contains("already exists", blockCollision.Message);
            Assert.Contains("already exists", regionCollision.Message);
        }

        [Fact]
        public void BlockDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("processor", "Processor", 0, 0)
                    .Callout("processor", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.BlockDiagram("Invalid", diagram => diagram
                    .Block("processor", "Processor", 0, 0)
                    .Callout("processor", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
