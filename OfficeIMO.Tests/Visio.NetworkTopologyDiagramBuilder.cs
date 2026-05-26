using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioNetworkTopologyDiagramBuilderTests {
        [Fact]
        public void NetworkTopologyDiagramBuilderCreatesAutomaticTopologyPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkTopologyDiagram("Branch Topology", topology => topology
                    .Theme(VisioStyleTheme.Technical())
                    .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
                    .Firewall("firewall", "Firewall")
                    .Switch("core", "Core Switch")
                    .Server("app", "App Server")
                    .Database("db", "Database")
                    .Workstation("pc1", "Finance PC")
                    .Workstation("pc2", "Support PC")
                    .Printer("printer", "Printer")
                    .Wireless("wifi", "Wi-Fi")
                    .Subnet("edge", "Edge", "internet", "firewall", "core")
                    .Subnet("server-zone", "Server Zone", "app", "db")
                    .Subnet("client-zone", "Client LAN", "pc1", "pc2", "printer", "wifi")
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Trunk("core", "app", "10Gb")
                    .Ethernet("app", "db")
                    .Ethernet("core", "pc1")
                    .Ethernet("core", "pc2")
                    .Ethernet("pc2", "printer")
                    .WirelessLink("core", "wifi", "wireless"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Branch Topology", page.Name);
            Assert.Equal(12, page.Shapes.Count);
            Assert.Equal(8, page.Connectors.Count);
            Assert.True(page.Shapes.Single(shape => shape.Id == "internet").PinX < page.Shapes.Single(shape => shape.Id == "firewall").PinX);
            Assert.True(page.Shapes.Single(shape => shape.Id == "firewall").PinX < page.Shapes.Single(shape => shape.Id == "core").PinX);
            Assert.True(page.Shapes.Single(shape => shape.Id == "app").PinX < page.Shapes.Single(shape => shape.Id == "db").PinX);
            Assert.Contains(page.Shapes, shape => shape.Id == "server-zone" && shape.IsBackgroundSurface);
            Assert.True(page.Shapes.Single(shape => shape.Id == "server-zone").Width > page.Shapes.Single(shape => shape.Id == "app").Width);
            Assert.Contains(page.Shapes, shape => shape.Id == "firewall" && shape.NameU == "Decision");
            Assert.Contains(page.Shapes, shape => shape.Id == "core" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "db" && shape.NameU == "Data");
            Assert.Contains(page.Shapes, shape => shape.Id == "wifi" && shape.NameU == "Circle");
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false
            }).Where(issue => issue.Severity >= VisioDiagramQualityIssueSeverity.Warning).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(12, loaded.Pages[0].Shapes.Count);
            Assert.Equal(8, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderAllowsMeshCycles() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .NetworkTopologyDiagram("Mesh", topology => topology
                    .Root("a", "Switch A", VisioNetworkNodeKind.Switch)
                    .Switch("b", "Switch B")
                    .Switch("c", "Switch C")
                    .Trunk("a", "b")
                    .Trunk("b", "c")
                    .Trunk("c", "a"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(3, page.Shapes.Count);
            Assert.Equal(3, page.Connectors.Count);
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderRejectsUnknownLinkEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Switch("core", "Core")
                    .Ethernet("core", "missing")));

            Assert.Contains("Unknown network node id", exception.Message);
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderRejectsUnknownZoneNodes() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid Zone", topology => topology
                    .Switch("core", "Core")
                    .Subnet("servers", "Servers", "core", "missing")));

            Assert.Contains("Unknown network node id", exception.Message);
        }
    }
}
