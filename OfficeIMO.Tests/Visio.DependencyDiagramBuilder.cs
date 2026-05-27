using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

using Xunit;

namespace OfficeIMO.Tests {
    public class VisioDependencyDiagramBuilderTests {
        [Fact]
        public void DependencyDiagramBuilderAutomaticallyLayersNodesAndRoutesDependencies() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .DependencyDiagram("Service Dependencies", diagram => diagram
                    .Theme(VisioStyleTheme.Fluent())
                    .PageSize(4, 3)
                    .External("user", "Users")
                    .Component("web", "Web App")
                    .Component("api", "API")
                    .Decision("policy", "Policy")
                    .Data("db", "Database")
                    .DependsOn("user", "web", "HTTPS")
                    .DependsOn("web", "api")
                    .ControlDependency("api", "policy", "Authorize")
                    .DataDependency("api", "db", "SQL"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Service Dependencies", page.Name);
            Assert.True(page.Width > 4);
            Assert.Equal(new[] { "user", "web", "api", "policy", "db" }, page.Shapes.Select(shape => shape.Id).ToArray());
            Assert.Equal("Circle", page.Shapes.Single(shape => shape.Id == "user").MasterNameU);
            Assert.Equal("Process", page.Shapes.Single(shape => shape.Id == "api").MasterNameU);
            Assert.Equal("Decision", page.Shapes.Single(shape => shape.Id == "policy").MasterNameU);
            Assert.Equal("Data", page.Shapes.Single(shape => shape.Id == "db").MasterNameU);
            Assert.True(page.Shapes.Single(shape => shape.Id == "user").PinX < page.Shapes.Single(shape => shape.Id == "web").PinX);
            Assert.True(page.Shapes.Single(shape => shape.Id == "web").PinX < page.Shapes.Single(shape => shape.Id == "api").PinX);
            Assert.True(page.Shapes.Single(shape => shape.Id == "api").PinX < page.Shapes.Single(shape => shape.Id == "db").PinX);
            Assert.Equal(4, page.Connectors.Count);
            Assert.Contains(page.Connectors, connector => connector.Label == "Authorize" && connector.LinePattern == 2);
            Assert.All(page.Connectors, connector => Assert.Equal(EndArrow.Triangle, connector.EndArrow));

            document.EnsureVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            });
            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(5, loaded.Pages[0].Shapes.Count);
            Assert.Equal(4, loaded.Pages[0].Connectors.Count);
            Assert.Contains(loaded.Pages[0].Connectors, connector => connector.Label == "HTTPS");
        }

        [Fact]
        public void DependencyDiagramBuilderRejectsCycles() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                document.DependencyDiagram("Cycle", diagram => diagram
                    .Component("a", "A")
                    .Component("b", "B")
                    .DependsOn("a", "b")
                    .DependsOn("b", "a")));

            Assert.Contains("cycle", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DependencyDiagramBuilderCanAddTitleWithoutOverlappingGraph() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .DependencyDiagram("Service Dependencies", diagram => diagram
                    .Title()
                    .External("users", "Users")
                    .Component("api", "API")
                    .Data("db", "Database")
                    .DependsOn("users", "api", "HTTPS")
                    .DataDependency("api", "db", "SQL"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            double highestNodeTop = page.Shapes
                .Where(shape => shape.Id != "title")
                .Max(shape => shape.PinY + shape.Height / 2D);
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Service Dependencies", title.Text);
            Assert.True(title.PinY - title.Height / 2D > highestNodeTop);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false,
                CheckConnectorLabelShapeOverlaps = false
            }).Where(issue => issue.Severity >= VisioDiagramQualityIssueSeverity.Warning).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void DependencyDiagramBuilderRejectsUnknownDependencyEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("known", "Known")
                    .DependsOn("known", "missing")));

            Assert.Contains("Unknown dependency node id", exception.Message);
        }

        [Fact]
        public void DependencyDiagramBuilderNormalizesDependencyEndpointIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .DependencyDiagram("Trimmed", diagram => diagram
                    .Component("api", "API")
                    .Data("db", "Database")
                    .DataDependency(" api ", " db ", "SQL"));

            VisioPage page = Assert.Single(document.Pages);
            VisioConnector connector = Assert.Single(page.Connectors);
            Assert.Equal("api", connector.From.Id);
            Assert.Equal("db", connector.To.Id);
            Assert.Equal("SQL", connector.Label);
        }

        [Fact]
        public void DependencyDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .DependencyDiagram("Annotated Dependencies", diagram => diagram
                    .Title()
                    .Component("api", "API")
                    .Decision("policy", "Policy")
                    .Data("db", "Database")
                    .ControlDependency("api", "policy", "Authorize")
                    .DataDependency("api", "db", "SQL")
                    .Callout(" policy ", "policy-note", "Policy gates this data path", 5.9, 5.6, options => {
                        options.Width = 2.5;
                        options.Height = 0.72;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "policy");
            Assert.Equal("policy-note", callout.Id);
            Assert.Equal("Policy gates this data path", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.5, callout.Width);
            Assert.Equal(0.72, callout.Height);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, callout));
            Assert.Same(target, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);
            Assert.Contains("Annotations", leader.LayerNames);
            Assert.Equal(leader.Id, callout.GetUserCellValue("OfficeIMO.CalloutLeaderId"));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void DependencyDiagramBuilderCanAutoPlaceSemanticCalloutsBesideNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .DependencyDiagram("Auto Annotated Dependencies", diagram => diagram
                    .Title()
                    .Component("api", "API")
                    .Decision("policy", "Policy")
                    .Data("db", "Database")
                    .ControlDependency("api", "policy", "Authorize")
                    .DataDependency("api", "db", "SQL")
                    .Callout("policy", "policy-note", "Policy gates this data path", VisioSide.Top, 0.4, options => {
                        options.Width = 2.5;
                        options.Height = 0.72;
                    })
                    .Callout("api", "Review client contracts", VisioSide.Left, 0.25));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape policy = Assert.Single(page.Shapes, shape => shape.Id == "policy");
            VisioShape api = Assert.Single(page.Shapes, shape => shape.Id == "api");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "policy-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "api-callout");

            Assert.True(explicitCallout.PinY > policy.PinY);
            Assert.Equal(policy.PinX, explicitCallout.PinX, 6);
            Assert.Equal(policy.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.5, explicitCallout.Width);
            Assert.True(generatedCallout.PinX < api.PinX);
            Assert.Equal(api.Id, generatedCallout.CalloutTargetId);
            VisioShapeBounds bounds = page.GetContentBounds();
            Assert.True(bounds.Left >= 0);
            Assert.True(bounds.Right <= page.Width);
            Assert.True(bounds.Bottom >= 0);
            Assert.True(bounds.Top <= page.Height);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(policy, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void DependencyDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .DependencyDiagram("Generated", diagram => diagram
                    .Component("api", "API")
                    .Callout("api", "First note", 3.5, 4.5)
                    .Callout("api", "Second note", 3.5, 3.6));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "api-callout", "api-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void DependencyDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException nodeFirst = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("title", "API")
                    .Title()));
            ArgumentException titleFirst = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Title()
                    .Component("title", "API")));

            Assert.Contains("already exists", nodeFirst.Message);
            Assert.Contains("already exists", titleFirst.Message);
        }

        [Fact]
        public void DependencyDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("api", "API")
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException nodeCollision = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("api", "API")
                    .Callout("api", "api", "Duplicate id", 4, 4)));
            ArgumentException titleCollision = Assert.Throws<ArgumentException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Title()
                    .Component("api", "API")
                    .Callout("api", "title", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown dependency node id", unknownTarget.Message);
            Assert.Contains("already exists", nodeCollision.Message);
            Assert.Contains("already exists", titleCollision.Message);
        }

        [Fact]
        public void DependencyDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("api", "API")
                    .Callout("api", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.DependencyDiagram("Invalid", diagram => diagram
                    .Component("api", "API")
                    .Callout("api", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("zero or greater", badGap.Message);
        }
    }
}
