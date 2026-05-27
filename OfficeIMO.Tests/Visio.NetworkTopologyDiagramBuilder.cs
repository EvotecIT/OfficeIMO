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
        public void NetworkTopologyDiagramBuilderCanAddTitleWithoutOverlappingZones() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkTopologyDiagram("Branch Topology", topology => topology
                    .Title()
                    .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
                    .Firewall("firewall", "Firewall")
                    .Switch("core", "Core Switch")
                    .Subnet("edge", "Edge", "internet", "firewall", "core")
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink"));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape title = Assert.Single(page.Shapes, shape => shape.Id == "title");
            VisioShape zone = Assert.Single(page.Shapes, shape => shape.Id == "edge");
            Assert.Equal("Text Box", title.NameU);
            Assert.Equal("Branch Topology", title.Text);
            Assert.True(title.PinY > zone.PinY);
            Assert.Empty(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckConnectorShapeIntersections = false
            }).Where(issue => issue.Severity >= VisioDiagramQualityIssueSeverity.Warning).Select(issue => issue.ToString()));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
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
        public void NetworkTopologyDiagramBuilderRejectsTitleIdCollisions() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Switch("title", "Core")
                    .Title()));

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderCanAddSemanticCallouts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkTopologyDiagram("Annotated Topology", topology => topology
                    .Title()
                    .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
                    .Firewall("firewall", "Firewall")
                    .Switch("core", "Core")
                    .Subnet("edge", "Edge", "internet", "firewall", "core")
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Callout(" firewall ", "firewall-note", "Inspect north-south traffic", 5.8, 5.7, options => {
                        options.Width = 2.55;
                        options.Height = 0.72;
                        options.RouteOffset = 0.1;
                    }));

            VisioPage page = Assert.Single(document.Pages);
            VisioShape callout = Assert.Single(page.Callouts());
            VisioShape target = Assert.Single(page.Shapes, shape => shape.Id == "firewall");
            Assert.Equal("firewall-note", callout.Id);
            Assert.Equal("Inspect north-south traffic", callout.Text);
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
        public void NetworkTopologyDiagramBuilderCanAutoPlaceSemanticCalloutsBesideNodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath)
                .NetworkTopologyDiagram("Auto Annotated Topology", topology => topology
                    .Title()
                    .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
                    .Firewall("firewall", "Firewall")
                    .Switch("core", "Core")
                    .Subnet("edge", "Edge", "internet", "firewall", "core")
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Callout("firewall", "firewall-note", "Inspect north-south traffic", VisioSide.Top, 0.4, options => {
                        options.Width = 2.55;
                        options.Height = 0.72;
                    })
                    .Callout("core", "Aggregation point", VisioSide.Right, 0.3));

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
        public void NetworkTopologyDiagramBuilderGeneratesUniqueCalloutIds() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"))
                .NetworkTopologyDiagram("Generated", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Callout("core", "First note", 3.5, 4.5)
                    .Callout("core", "Second note", 3.5, 3.6));

            VisioPage page = Assert.Single(document.Pages);
            Assert.Equal(new[] { "core-callout", "core-callout-2" }, page.Callouts().Select(shape => shape.Id).ToArray());
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderRejectsCalloutIdCollisionsAndUnknownTargets() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentException unknownTarget = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Callout("missing", "note", "No target", 4, 4)));
            ArgumentException nodeCollision = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Callout("core", "core", "Duplicate id", 4, 4)));
            ArgumentException zoneCollision = Assert.Throws<ArgumentException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Subnet("edge", "Edge", "core")
                    .Callout("core", "edge", "Duplicate id", 4, 4)));

            Assert.Contains("Unknown network node id", unknownTarget.Message);
            Assert.Contains("already exists", nodeCollision.Message);
            Assert.Contains("already exists", zoneCollision.Message);
        }

        [Fact]
        public void NetworkTopologyDiagramBuilderRejectsAutoCalloutPlacementIssues() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));

            ArgumentOutOfRangeException autoPlacement = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Callout("core", "Invalid", VisioSide.Auto)));
            ArgumentOutOfRangeException badGap = Assert.Throws<ArgumentOutOfRangeException>(() =>
                document.NetworkTopologyDiagram("Invalid", topology => topology
                    .Root("core", "Core", VisioNetworkNodeKind.Switch)
                    .Callout("core", "Invalid", VisioSide.Right, double.NaN)));

            Assert.Contains("Placement must be", autoPlacement.Message);
            Assert.Contains("zero or greater", badGap.Message);
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
