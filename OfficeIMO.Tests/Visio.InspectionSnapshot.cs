using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioInspectionSnapshotTests {
        [Fact]
        public void InspectionSnapshotCapturesStableStructureAndSemanticData() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            document.Title = "Inspection sample";
            document.Author = "OfficeIMO";
            document.UseMastersByDefault = true;
            VisioPage page = document.AddPage("Topology", 8, 5);
            VisioShape gateway = page.AddRectangle(1.5, 3.5, 1.4, 0.7, "Gateway");
            gateway.NameU = "Rectangle";
            gateway.SetShapeData("Owner", "Platform", "Owner", VisioShapeDataType.String);
            gateway.SetUserCell(VisioSemanticUserCells.Kind, "Gateway", "STR", prompt: "semantic kind");
            gateway.Data["Tier"] = "Edge";
            VisioShape service = page.AddRectangle(5.5, 3.5, 1.4, 0.7, "Service");
            VisioConnector connector = page.AddConnector(gateway, service, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .PlaceLabel(0.5);
            connector.Label = "HTTPS";
            connector.SetShapeData("Protocol", "TLS", "Protocol", VisioShapeDataType.String);
            page.AddCallout(service, "service-note", "SLO: 99.9%", VisioSide.Bottom);

            VisioInspectionSnapshot snapshot = document.CreateInspectionSnapshot();
            string text = snapshot.ToText();

            Assert.Equal("Inspection sample", snapshot.Title);
            VisioInspectionPageSnapshot pageSnapshot = Assert.Single(snapshot.Pages);
            Assert.True(snapshot.ShapeCount >= 3);
            Assert.Single(pageSnapshot.Connectors, item => item.Label == "HTTPS");
            VisioInspectionShapeSnapshot gatewaySnapshot = Assert.Single(pageSnapshot.Shapes, shape => shape.Id == gateway.Id);
            Assert.Equal("Gateway", gatewaySnapshot.UserCells.Single(cell => cell.Name == VisioSemanticUserCells.Kind).Value);
            Assert.Equal("Platform", gatewaySnapshot.ShapeData.Single(row => row.Name == "Owner").Value);
            Assert.Equal("Edge", gatewaySnapshot.Data["Tier"]);
            Assert.Single(gatewaySnapshot.ConnectionPoints);
            Assert.Equal(1.4, gatewaySnapshot.ConnectionPoints[0].X, 6);
            Assert.Equal(0.35, gatewaySnapshot.ConnectionPoints[0].Y, 6);
            Assert.Contains("document.title=Inspection sample", text, StringComparison.Ordinal);
            Assert.Contains("page[Topology].shape[" + gateway.Id + "].connectionPointCount=1", text, StringComparison.Ordinal);
            Assert.Contains("page[Topology].shape[" + gateway.Id + "].connectionPoint[0].dirX=-1", text, StringComparison.Ordinal);
            Assert.Contains("page[Topology].shape[" + gateway.Id + "].shapeData[Owner].value=Platform", text, StringComparison.Ordinal);
            Assert.Contains("page[Topology].connector[" + connector.Id + "].label=HTTPS", text, StringComparison.Ordinal);
            Assert.Contains("isCallout=true", text, StringComparison.Ordinal);
        }

        [Fact]
        public void InspectionDiffReportsChangedAndAddedPaths() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Diff", 6, 4);
            VisioShape start = page.AddRectangle(1, 2, 1, 0.5, "Start");

            VisioInspectionSnapshot before = document.CreateInspectionSnapshot();

            start.Text = "Renamed";
            page.AddRectangle(3, 2, 1, 0.5, "Added");
            VisioInspectionSnapshot after = document.CreateInspectionSnapshot();

            VisioInspectionDiff diff = before.Diff(after);

            Assert.True(diff.HasDifferences);
            Assert.Contains(diff.Differences, difference =>
                difference.Kind == VisioInspectionDifferenceKind.Changed &&
                difference.Path == "page[Diff].shape[" + start.Id + "].text" &&
                difference.Expected == "Start" &&
                difference.Actual == "Renamed");
            Assert.Contains(diff.Differences, difference =>
                difference.Kind == VisioInspectionDifferenceKind.Added &&
                difference.Path.EndsWith(".text", StringComparison.Ordinal) &&
                difference.Actual == "Added");
        }

        [Fact]
        public void InspectionSnapshotRoundTripsLoadedShapeData() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("RoundTrip", 6, 4);
            VisioShape node = page.AddRectangle(1, 2, 1, 0.5, "Node");
            node.SetShapeData("Criticality", "Tier 0", "Criticality", VisioShapeDataType.String);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioInspectionSnapshot snapshot = loaded.CreateInspectionSnapshot();

            VisioInspectionShapeSnapshot loadedNode = Assert.Single(snapshot.Pages[0].Shapes, shape => shape.Text == "Node");
            Assert.Equal("Tier 0", loadedNode.ShapeData.Single(row => row.Name == "Criticality").Value);
        }
    }
}
