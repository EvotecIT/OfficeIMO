using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioNetworkDiagramBuilderTests {
        [Fact]
        public void NetworkDiagramBuilderCreatesStyledNetworkPage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkDiagram("Branch Network", network => network
                    .Theme(VisioStyleTheme.Technical())
                    .Zone("perimeter", "Perimeter", 0, 0, 3, 1)
                    .Zone("servers", "Server Zone", 3, 0, 3, 1)
                    .Zone("clients", "Client LAN", 1, 2, 5, 1)
                    .Internet("internet", "Internet", 0, 0)
                    .Firewall("firewall", "Firewall", 1, 0)
                    .Switch("core", "Core Switch", 2, 0)
                    .Server("app", "App Server", 3, 0)
                    .Database("db", "Database", 4, 0)
                    .Storage("backup", "Backup NAS", 5, 0)
                    .Workstation("pc1", "Finance PC", 1, 2)
                    .Workstation("pc2", "Support PC", 2, 2)
                    .Printer("printer", "Printer", 3, 2)
                    .Wireless("wifi", "Wi-Fi", 4, 2)
                    .Legend("legend", "solid: data\ndashed: mgmt", 5, 2)
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Trunk("core", "app", "10Gb")
                    .Ethernet("app", "db")
                    .Ethernet("db", "backup")
                    .Ethernet("core", "pc2")
                    .Ethernet("pc1", "pc2")
                    .Ethernet("pc2", "printer")
                    .WirelessLink("printer", "wifi", "wireless"));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal("Branch Network", page.Name);
            Assert.Equal(14, page.Shapes.Count);
            Assert.Equal(9, page.Connectors.Count);
            Assert.Contains(page.Shapes, shape => shape.Id == "firewall" && shape.NameU == "Decision");
            Assert.Contains(page.Shapes, shape => shape.Id == "core" && shape.NameU == "Rectangle");
            Assert.Contains(page.Shapes, shape => shape.Id == "db" && shape.NameU == "Data");
            Assert.Contains(page.Shapes, shape => shape.Id == "wifi" && shape.NameU == "Circle");
            Assert.All(page.Connectors, connector => Assert.NotEmpty(connector.Waypoints));
            Assert.Empty(page.AnalyzeVisualQuality().Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Equal(14, loaded.Pages[0].Shapes.Count);
            Assert.Equal(9, loaded.Pages[0].Connectors.Count);
        }

        [Fact]
        public void NetworkStencilCatalogExposesCommonNetworkShapes() {
            Assert.Equal("Network", VisioStencils.Network.Name);
            Assert.Equal("Switch", VisioStencils.Network.Get("lan").Name);
            Assert.Equal("Firewall", VisioStencils.Network.Get("security").Name);
            Assert.Equal("Wireless AP", VisioStencils.All.Get("net.wireless").Name);
        }

        [Fact]
        public void NetworkDiagramBuilderCanAddTitleWithoutOverlappingZones() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkDiagram("Branch Network", network => network
                    .Title()
                    .Zone("perimeter", "Perimeter", 0, 0, 3, 1)
                    .Internet("internet", "Internet", 0, 0)
                    .Firewall("firewall", "Firewall", 1, 0)
                    .Switch("core", "Core Switch", 2, 0)
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape zone = Assert.Single(page.Shapes, shape => shape.Id == "perimeter");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Branch Network", title.Text);
            Assert.True(title.PinY > zone.PinY);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false
            }).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void NetworkDiagramBuilderRejectsUnknownLinkEndpoints() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("core", "Core", 0, 0)
                    .Ethernet("core", "missing")));

            Assert.Contains("Unknown network node id", exception.Message);
        }

        [Fact]
        public void NetworkDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("title", "Core", 0, 0)
                    .Title()));

            Assert.Contains("already exists", exception.Message);
        }
    }
}
