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
        public void NetworkDiagramBuilderNormalizesLinkEndpointIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .NetworkDiagram("Trimmed", network => network
                    .Switch("core", "Core", 0, 0)
                    .Server("app", "App", 1, 0)
                    .Ethernet(" core ", " app ", "LAN"));

            VisioPage page = Assert.Single(document.Pages);
            VisioConnector connector = Assert.Single(page.Connectors);
            Assert.Equal("core", connector.From.Id);
            Assert.Equal("app", connector.To.Id);
            Assert.Equal("LAN", connector.Label);
        }

        [Fact]
        public void NetworkDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkDiagram("Annotated Network", network => network
                    .Title()
                    .Zone("edge", "Edge", 0, 0, 3, 1)
                    .Internet("internet", "Internet", 0, 0)
                    .Firewall("firewall", "Firewall", 1, 0)
                    .Switch("core", "Core", 2, 0)
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Callout(" firewall ", "firewall-note", "Inbound traffic inspection", 5.6, 6.2, options => {
                        options.Width = 2.55;
                        options.Height = 0.72;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "firewall");
            Assert.Equal("firewall-note", callout.Id);
            Assert.Equal("Inbound traffic inspection", callout.Text);
            Assert.Equal(target.Id, callout.CalloutTargetId);
            Assert.Contains("Annotations", callout.LayerNames);
            Assert.Equal(2.55, callout.Width);
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
        public void NetworkDiagramBuilderCanAutoPlaceSemanticCalloutsBesideNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkDiagram("Auto Annotated Network", network => network
                    .Title()
                    .Zone("edge", "Edge", 0, 0, 3, 1)
                    .Internet("internet", "Internet", 0, 0)
                    .Firewall("firewall", "Firewall", 1, 0)
                    .Switch("core", "Core", 2, 0)
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Callout("firewall", "firewall-note", "Inspect and log inbound traffic", VisioSide.Top, 0.4, options => {
                        options.Width = 2.55;
                        options.Height = 0.72;
                    })
                    .Callout("core", "Redundant uplink target", VisioSide.Right, 0.3));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape firewall = Assert.Single(page.Shapes, shape => shape.Id == "firewall");
            VisioShape core = Assert.Single(page.Shapes, shape => shape.Id == "core");
            VisioShape explicitCallout = Assert.Single(page.Callouts(), shape => shape.Id == "firewall-note");
            VisioShape generatedCallout = Assert.Single(page.Callouts(), shape => shape.Id == "core-callout");

            Assert.True(explicitCallout.PinY > firewall.PinY);
            Assert.Equal(firewall.PinX, explicitCallout.PinX, 6);
            Assert.Equal(firewall.Id, explicitCallout.CalloutTargetId);
            Assert.Equal(2.55, explicitCallout.Width);
            Assert.True(generatedCallout.PinX > core.PinX);
            Assert.Equal(core.Id, generatedCallout.CalloutTargetId);

            VisioConnector leader = Assert.Single(page.Connectors, connector => ReferenceEquals(connector.From, explicitCallout));
            Assert.Same(firewall, leader.To);
            Assert.Equal(EndArrow.None, leader.EndArrow);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void NetworkDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .NetworkDiagram("Generated", network => network
                    .Switch("core", "Core", 0, 0)
                    .Callout("core", "First note", 3.5, 4.5)
                    .Callout("core", "Second note", 3.5, 3.6));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "core-callout", "core-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
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

        [Fact]
        public void NetworkDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("core", "Core", 0, 0)
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException nodeCollision = Assert.Throws<ArgumentException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("core", "Core", 0, 0)
                    .Callout("core", "core", "Duplicate id", 4, 4)));
            ArgumentException zoneCollision = Assert.Throws<ArgumentException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Zone("edge", "Edge", 0, 0, 1, 1)
                    .Switch("core", "Core", 0, 0)
                    .Callout("core", "edge", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown network node id", unknownTarget.Message);
            Assert.Contains("already exists", nodeCollision.Message);
            Assert.Contains("already exists", zoneCollision.Message);
        }

        [Fact]
        public void NetworkDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("core", "Core", 0, 0)
                    .Callout("core", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.NetworkDiagram("Invalid", network => network
                    .Switch("core", "Core", 0, 0)
                    .Callout("core", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("finite non-negative", badGap.Message);
        }
    }
}
