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
        public void FlowchartBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Annotated", flow => flow
                    .Start("start", "Start")
                    .Step("review", "Review")
                    .Decision("agreement", "Agreement?")
                    .End("done", "Done")
                    .Callout(" agreement ", "agreement-note", "Escalate if rejected", 6.2, 5.7, options => {
                        options.Width = 2.35;
                        options.Height = 0.7;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "agreement");
            Assert.Equal("agreement-note", callout.Id);
            Assert.Equal("Escalate if rejected", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.35, callout.Width);
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
        public void FlowchartBuilderCanAutoPlaceSemanticCalloutsBesideNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .Flowchart("Auto Annotated", flow => flow
                    .Start("start", "Start")
                    .Step("review", "Review")
                    .Decision("agreement", "Agreement?")
                    .End("done", "Done")
                    .Callout("agreement", "agreement-note", "Escalate if rejected", VisioSide.Right, 0.4, options => {
                        options.Width = 2.35;
                        options.Height = 0.7;
                    })
                    .Callout("review", "Check completeness", VisioSide.Left, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "agreement");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "agreement-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "review-callout");

            Assert.True(explicitCallout.PinX > target.PinX);
            Assert.Equal(target.PinY, explicitCallout.PinY, 6);
            Assert.Equal(target.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.35, explicitCallout.Width);
            Assert.True(generatedCallout.PinX < page.FindShapeById("review")!.PinX);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void FlowchartBuilderNormalizesNodeAndConnectorIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .Flowchart("Trimmed", flow => flow
                    .Start(" start ", "Start")
                    .Step(" review ", "Review")
                    .End(" done ", "Done")
                    .Branch(" review ", "retry", " start "));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Contains(page.Shapes, shape => shape.Id == "start");
            Assert.Contains(page.Shapes, shape => shape.Id == "review");
            Assert.Contains(page.Shapes, shape => shape.Id == "done");
            VisioConnector connector = Assert.Single(page.Connectors, item => item.Label == "retry");
            Assert.Equal("review", connector.From.Id);
            Assert.Equal("start", connector.To.Id);
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

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void FlowchartBuilderRejectsCalloutIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException nodeCollision = Assert.Throws<ArgumentException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Start("start", "Start")
                    .Callout("start", "start", "Duplicate id", 4, 4)));
            ArgumentException titleCollision = Assert.Throws<ArgumentException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Title()
                    .Start("start", "Start")
                    .Callout("start", "title", "Duplicate id", 4, 4)));

            Assert.Contains("already exists", nodeCollision.Message);
            Assert.Contains("already exists", titleCollision.Message);
        }

        [Fact]
        public void FlowchartBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Start("start", "Start")
                    .Callout("start", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.Flowchart("Invalid", flow => flow
                    .Start("start", "Start")
                    .Callout("start", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
