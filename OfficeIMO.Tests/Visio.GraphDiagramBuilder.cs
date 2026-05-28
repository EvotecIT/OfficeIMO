using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Tests {
    public class VisioGraphDiagramBuilderTests {
        [Fact]
        public void GraphDiagramBuilderHandlesCyclesDisconnectedComponentsAndStencilNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilShape serverStencil = VisioStencils.Network.Get("server");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Runtime Graph", graph => graph
                    .Title()
                    .Theme(VisioStyleTheme.Fluent())
                    .Root("users", "Users", VisioGraphNodeKind.External)
                    .StencilNode("web", "Web", serverStencil)
                    .Node("api", "API", VisioGraphNodeKind.Process)
                    .Node("policy", "Policy", VisioGraphNodeKind.Decision)
                    .Node("db", "Database", VisioGraphNodeKind.Data)
                    .Node("batch", "Batch", VisioGraphNodeKind.Emphasis)
                    .Zone("online", "Online path", "users", "web", "api", "policy", "db")
                    .Zone("offline", "Offline path", "batch", "db")
                    .Edge("users", "web", "HTTPS")
                    .Edge("web", "api")
                    .ControlEdge("api", "policy", "authorize")
                    .DataEdge("api", "db", "read/write")
                    .DataEdge("db", "api", "cache")
                    .Relationship("batch", "db", "nightly"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Runtime Graph", page.Name);
            Assert.Contains(page.Shapes, shape => shape.Id == "online" && shape.IsBackgroundSurface);
            Assert.Contains(page.Shapes, shape => shape.Id == "offline" && shape.IsBackgroundSurface);
            Assert.Contains(page.Shapes, shape => shape.Id == "online-label" && shape.Text == "Online path");
            Assert.Contains(page.Shapes, shape => shape.Id == "web-label" && shape.Text == "Web");
            Assert.True(string.IsNullOrEmpty(page.Shapes.Single(shape => shape.Id == "web").Text));
            Assert.Equal(serverStencil.MasterNameU, page.Shapes.Single(shape => shape.Id == "web").MasterNameU);
            Assert.True(page.Shapes.Single(shape => shape.Id == "users").PinX < page.Shapes.Single(shape => shape.Id == "web").PinX);
            Assert.True(page.Shapes.Single(shape => shape.Id == "web").PinX < page.Shapes.Single(shape => shape.Id == "api").PinX);
            Assert.Equal(6, page.Connectors.Count);
            Assert.Contains(page.Connectors, connector => connector.Label == "authorize" && connector.LinePattern == 2);
            Assert.Contains(page.Connectors, connector => connector.Label == "nightly" && connector.EndArrow == EndArrow.None);
            Assert.All(page.Connectors.Where(connector => !string.IsNullOrWhiteSpace(connector.Label)), connector => Assert.NotNull(connector.LabelPlacement));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(12, loaded.Pages[0].Shapes.Count);
            Assert.Equal(6, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void GraphDiagramBuilderSupportsRadialLayoutForCyclicGraphs() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Cyclic Service Map", graph => graph
                    .Layout(VisioGraphLayout.Radial)
                    .Root("api", "API")
                    .Node("auth", "Auth")
                    .Node("queue", "Queue")
                    .Node("worker", "Worker")
                    .Node("store", "Store", VisioGraphNodeKind.Data)
                    .Edge("api", "auth", "token")
                    .Edge("api", "queue", "publish")
                    .Edge("queue", "worker", "consume")
                    .DataEdge("worker", "store", "write")
                    .DataEdge("store", "api", "read"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape api = page.Shapes.Single(shape => shape.Id == "api");
            Assert.Equal(5, page.Shapes.Count);
            Assert.Equal(5, page.Connectors.Count);
            Assert.All(page.Connectors, connector => Assert.Empty(connector.Waypoints));
            Assert.Contains(page.Shapes.Where(shape => shape.Id != "api"), shape => Math.Abs(shape.PinX - api.PinX) > 0.1D || Math.Abs(shape.PinY - api.PinY) > 0.1D);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderAvoidsGeneratedStencilCaptionIdCollisions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioStencilShape serverStencil = VisioStencils.Network.Get("server");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Collision Graph", graph => graph
                    .StencilNode("web", "Web", serverStencil)
                    .Node("web-label", "Existing label id", VisioGraphNodeKind.Process)
                    .Edge("web", "web-label"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Contains(page.Shapes, shape => shape.Id == "web");
            Assert.Contains(page.Shapes, shape => shape.Id == "web-label" && shape.Text == "Existing label id");
            Assert.Contains(page.Shapes, shape => shape.Id == "web-label-2" && shape.Text == "Web");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderAvoidsNamedNumericEdgeIdCollisions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Numeric Edge Ids", graph => graph
                    .Node("source", "Source")
                    .Node("middle", "Middle")
                    .Node("target", "Target")
                    .Edge("source", "middle")
                    .DataEdge("1", "middle", "target", "uses"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(page.Connectors.Count, page.Connectors.Select(connector => connector.Id).Distinct().Count());
            Assert.Contains(page.Connectors, connector => connector.Id == "1" && connector.Label == "uses");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderCanSelectStencilNodesFromCatalogQueries() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Catalog Query Graph", graph => graph
                    .StencilNode("client", "Client", VisioStencils.Network, "missing", "access-point")
                    .StencilNode("service", "Service", VisioStencils.Flowchart, "process")
                    .Edge("client", "service", "calls"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Circle", page.Shapes.Single(shape => shape.Id == "client").MasterNameU);
            Assert.Equal("Process", page.Shapes.Single(shape => shape.Id == "service").MasterNameU);
            Assert.Contains(page.Shapes, shape => shape.Id == "client-label" && shape.Text == "Client");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderFitsStencilCaptionOverflow() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Small Stencil Graph", graph => graph
                    .PageSize(2.4, 1.6)
                    .StencilNode("api", "API", VisioStencils.Flowchart, "process"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.True(page.Height >= 2.94D);
            Assert.Contains(page.Shapes, shape => shape.Id == "api-label" && shape.Text == "API");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderCanAttachShapeDataAndHyperlinksToNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Metadata Graph", graph => graph
                    .Node("api", "API", VisioGraphNodeKind.Emphasis)
                    .Node("database", "Database", VisioGraphNodeKind.Data)
                    .NodeShapeData("api", "Owner", "Platform", "Owner", VisioShapeDataType.String, "Owning team")
                    .NodeShapeData("api", "Tier", "Public", "Service tier", VisioShapeDataType.String)
                    .NodeShapeData("database", "Classification", "Confidential", "Data classification", VisioShapeDataType.String)
                    .NodeHyperlink("api", "https://example.org/runbook", "Runbook")
                    .NodeHyperlink("database", new Uri("https://example.org/data-catalog"), "Data catalog")
                    .DataEdge("api-reads-database", "api", "database", "reads")
                    .EdgeShapeData("api-reads-database", "Protocol", "HTTPS", "Protocol", VisioShapeDataType.String)
                    .EdgeShapeData("api-reads-database", "Port", "443", "Port", VisioShapeDataType.Number)
                    .EdgeHyperlink("api-reads-database", "https://example.org/openapi.json", "API contract"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape api = page.Shapes.Single(shape => shape.Id == "api");
            VisioShape database = page.Shapes.Single(shape => shape.Id == "database");
            VisioConnector connector = page.Connectors.Single(edge => edge.Id == "api-reads-database");
            Assert.Equal("Platform", api.GetShapeDataValue("Owner"));
            Assert.Equal("Public", api.GetShapeDataValue("Tier"));
            Assert.Equal("Confidential", database.GetShapeDataValue("Classification"));
            Assert.Contains(api.ShapeData, row => row.Name == "Owner" && row.Label == "Owner" && row.Prompt == "Owning team");
            Assert.Contains(api.Hyperlinks, hyperlink => hyperlink.Address == "https://example.org/runbook" && hyperlink.Description == "Runbook");
            Assert.Contains(database.Hyperlinks, hyperlink => hyperlink.Address == "https://example.org/data-catalog" && hyperlink.Description == "Data catalog");
            Assert.Equal("HTTPS", connector.GetShapeDataValue("Protocol"));
            Assert.Equal("443", connector.GetShapeDataValue("Port"));
            Assert.Contains(connector.ShapeData, row => row.Name == "Protocol" && row.Label == "Protocol" && row.Type == VisioShapeDataType.String);
            Assert.Contains(connector.Hyperlinks, hyperlink => hyperlink.Address == "https://example.org/openapi.json" && hyperlink.Description == "API contract");

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape loadedApi = loaded.Pages[0].Shapes.Single(shape => shape.Id == "api");
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single(edge => edge.Id == "api-reads-database");
            Assert.Equal("Platform", loadedApi.GetShapeDataValue("Owner"));
            Assert.Contains(loadedApi.Hyperlinks, hyperlink => hyperlink.Address == "https://example.org/runbook" && hyperlink.Description == "Runbook");
            Assert.Equal("HTTPS", loadedConnector.GetShapeDataValue("Protocol"));
            Assert.Equal("443", loadedConnector.GetShapeDataValue("Port"));
            Assert.Contains(loadedConnector.Hyperlinks, hyperlink => hyperlink.Address == "https://example.org/openapi.json" && hyperlink.Description == "API contract");
        }

        [Fact]
        public void GraphDiagramBuilderCanOverrideNodeAndNamedEdgeStyles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            Color nodeFill = Color.FromRgb(245, 104, 85);
            Color nodeLine = Color.FromRgb(160, 44, 34);
            Color edgeLine = Color.FromRgb(112, 48, 160);

            VisioDocument document = VisioDocument.Create(filePath)
                .GraphDiagram("Styled Graph", graph => graph
                    .Theme(VisioStyleTheme.Technical())
                    .Node("api", "API")
                    .Node("database", "Database", VisioGraphNodeKind.Data)
                    .NodeStyle("api", style => {
                        style.FillColor = nodeFill;
                        style.LineColor = nodeLine;
                        style.LineWeight = 0.031D;
                    })
                    .DataEdge("api-reads-database", "api", "database", "reads")
                    .EdgeStyle("api-reads-database", style => {
                        style.LineColor = edgeLine;
                        style.LineWeight = 0.033D;
                        style.LinePattern = 2;
                        style.EndArrow = EndArrow.Arrow;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape api = page.Shapes.Single(shape => shape.Id == "api");
            VisioConnector connector = page.Connectors.Single(edge => edge.Id == "api-reads-database");
            Assert.Equal(nodeFill, api.FillColor);
            Assert.Equal(nodeLine, api.LineColor);
            Assert.Equal(0.031D, api.LineWeight, 3);
            Assert.Equal(edgeLine, connector.LineColor);
            Assert.Equal(0.033D, connector.LineWeight, 3);
            Assert.Equal(2, connector.LinePattern);
            Assert.Equal(EndArrow.Arrow, connector.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void GraphDiagramBuilderRejectsUnknownEdgeEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.GraphDiagram("Invalid", graph => graph
                    .Node("known", "Known")
                    .Edge("known", "missing")));

            Assert.Contains("Unknown graph node id", exception.Message);
        }

        [Fact]
        public void GraphDiagramBuilderRejectsUnknownNamedEdgesAndDuplicateEdgeIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknown = Assert.Throws<ArgumentException>(() =>
                document.GraphDiagram("Unknown Edge", graph => graph
                    .Node("api", "API")
                    .Node("database", "Database")
                    .DataEdge("api-reads-database", "api", "database", "reads")
                    .EdgeHyperlink("missing-edge", "https://example.org")));

            Assert.Contains("Unknown graph edge id", unknown.Message);

            ArgumentException duplicate = Assert.Throws<ArgumentException>(() =>
                VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                    .GraphDiagram("Duplicate Edge", graph => graph
                        .Node("api", "API")
                        .Node("database", "Database")
                        .DataEdge("api", "api", "database", "reads")));

            Assert.Contains("already exists", duplicate.Message);
        }
    }
}
