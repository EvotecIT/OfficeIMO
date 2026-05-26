using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioFlowchartBuilderTests {
        [Fact]
        public void FlowchartBuilderCreatesStyledSemanticPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Buying", flow => flow
                    .Start("start", "Start")
                    .Step("consult", "Consult")
                    .Decision("agreement", "Agreement?")
                    .End("close", "Close")
                    .Branch("agreement", "No", "consult"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Buying", page.Name);
            Assert.Equal(new[] { "start", "consult", "agreement", "close" }, page.Shapes.Select(shape => shape.Id).ToArray());
            Assert.Equal(new[] { "Ellipse", "Process", "Decision", "Ellipse" }, page.Shapes.Select(shape => shape.NameU).ToArray());
            Assert.Equal(new[] { "Ellipse", "Process", "Decision", "Ellipse" }, page.Shapes.Select(shape => shape.Master?.NameU).ToArray());
            Assert.Equal(4, page.Connectors.Count);
            Assert.Contains(page.Connectors, connector => connector.Label == "No");
            Assert.All(page.Connectors, connector => Assert.Equal(EndArrow.Triangle, connector.EndArrow));
            Assert.All(page.Connectors, connector => Assert.Equal(ConnectorKind.RightAngle, connector.Kind));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(4, loaded.Pages[0].Shapes.Count);
            Assert.Equal(4, loaded.Pages[0].Connectors.Count);
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "No");
        }

        [Fact]
        public void TwoColumnContinuationSplitsAtContinuationMarker() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .Flowchart("Two Column", flow => flow
                    .Layout(VisioFlowchartLayout.TwoColumnContinuation)
                    .Start("start", "Start")
                    .Step("left", "Left")
                    .OffPage("jump", "A")
                    .Continue("resume", "A")
                    .Step("right", "Right")
                    .End("end", "End"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(6, page.Shapes.Count);
            Assert.True(page.Shapes[0].PinX < page.Shapes[3].PinX);
            Assert.True(page.Shapes[1].PinX < page.Shapes[4].PinX);
            Assert.Equal("Off-page reference", page.Shapes[2].NameU);
            Assert.Equal("Circle", page.Shapes[3].NameU);
            Assert.DoesNotContain(page.Connectors, connector => connector.From.Id == "jump" && connector.To.Id == "resume");
        }

        [Fact]
        public void FlowchartBuilderRoutesComplexBranchesAroundShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Complex Branches", flow => flow
                    .Layout(VisioFlowchartLayout.TwoColumnContinuation)
                    .RouteBranches(laneSpacing: 0.5)
                    .Start("start", "Start")
                    .Step("intake", "Intake")
                    .Step("review", "Review")
                    .Step("market", "Choose option")
                    .OffPage("jump", "A")
                    .Continue("resume", "A")
                    .Step("offer", "Make offer")
                    .Decision("agreement", "Agreement?")
                    .Step("contract", "Contract")
                    .End("done", "Done")
                    .Branch("agreement", "No", "market"));

            VisioPage page = Assert.Single(document.Pages);
            VisioConnector branch = Assert.Single(page.Connectors, connector => connector.Label == "No");
            Assert.NotEmpty(branch.Waypoints);
            Assert.NotNull(branch.LabelPlacement);
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void FlowchartBuilderCanAddTitleWithoutOverlappingNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Property buying Flowchart", flow => flow
                    .Title()
                    .Start("start", "Start")
                    .Step("consult", "Consult")
                    .Decision("agreement", "Agreement?")
                    .End("close", "Close"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape start = Assert.Single(page.Shapes, shape => shape.Id == "start");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Property buying Flowchart", title.Text);
            Assert.True(title.PinY > start.PinY);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabels = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void FlowchartBuilderRejectsDuplicateNodeIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Start("same", "Start")
                    .Step("same", "Step")));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void FlowchartBuilderRejectsTitleNodeIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Title(id: "start")
                    .Start("start", "Start")));

            Assert.Contains("title with id", exception.Message);
        }
    }
}
