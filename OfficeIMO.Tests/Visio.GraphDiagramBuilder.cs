using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

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
        public void GraphDiagramBuilderRejectsUnknownEdgeEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.GraphDiagram("Invalid", graph => graph
                    .Node("known", "Known")
                    .Edge("known", "missing")));

            Assert.Contains("Unknown graph node id", exception.Message);
        }
    }
}
