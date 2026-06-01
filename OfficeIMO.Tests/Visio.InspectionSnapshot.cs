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
            document.Theme = new VisioTheme { Name = "Premium Blue" };
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
            Assert.Equal("Premium Blue", snapshot.ThemeType);
            VisioInspectionPageSnapshot pageSnapshot = Assert.Single(snapshot.Pages);
            Assert.True(snapshot.ShapeCount >= 3);
            VisioInspectionConnectorSnapshot connectorSnapshot = Assert.Single(pageSnapshot.Connectors, item => item.Label == "HTTPS");
            Assert.True(connectorSnapshot.HasLabelPlacement);
            Assert.Equal(0.5, connectorSnapshot.LabelPosition);
            Assert.Equal(0, connectorSnapshot.LabelOffsetX);
            Assert.Equal(0, connectorSnapshot.LabelOffsetY);
            Assert.Equal(1.25, connectorSnapshot.LabelWidth);
            Assert.Equal(0.3, connectorSnapshot.LabelHeight);
            Assert.Equal(3.5, connectorSnapshot.LabelResolvedPinX);
            Assert.Equal(3.5, connectorSnapshot.LabelResolvedPinY);
            Assert.Equal(0.625, connectorSnapshot.LabelLocPinX);
            Assert.Equal(0.15, connectorSnapshot.LabelLocPinY);
            VisioInspectionShapeSnapshot gatewaySnapshot = Assert.Single(pageSnapshot.Shapes, shape => shape.Id == gateway.Id);
            Assert.Equal("Gateway", gatewaySnapshot.UserCells.Single(cell => cell.Name == VisioSemanticUserCells.Kind).Value);
            Assert.Equal("Platform", gatewaySnapshot.ShapeData.Single(row => row.Name == "Owner").Value);
            Assert.Equal("Edge", gatewaySnapshot.Data["Tier"]);
            Assert.Single(gatewaySnapshot.ConnectionPoints);
            Assert.Equal(1.4, gatewaySnapshot.ConnectionPoints[0].X, 6);
            Assert.Equal(0.35, gatewaySnapshot.ConnectionPoints[0].Y, 6);
            Assert.Contains("document.title=Inspection sample", text, StringComparison.Ordinal);
            Assert.Contains("document.theme=Premium Blue", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].shape[" + gateway.Id + "].connectionPointCount=1", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].shape[" + gateway.Id + "].connectionPoint[0].dirX=-1", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].shape[" + gateway.Id + "].shapeData[Owner].value=Platform", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].connector[" + connector.Id + "].label=HTTPS", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].connector[" + connector.Id + "].labelPosition=0.5", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].connector[" + connector.Id + "].labelResolvedPinX=3.5", text, StringComparison.Ordinal);
            Assert.Contains("page[" + page.Id + ":Topology].connector[" + connector.Id + "].labelWidth=1.25", text, StringComparison.Ordinal);
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
                difference.Path == "page[" + page.Id + ":Diff].shape[" + start.Id + "].text" &&
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

        [Fact]
        public void InspectionSnapshotEscapesLineValuesAndDelimiterKeys() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Escapes", 6, 4);
            VisioShape node = page.AddRectangle(1, 2, 1, 0.5, "Line 1\nLine 2");
            node.Data["owner=team\nprimary"] = "value=one\nvalue two";

            VisioInspectionSnapshot before = document.CreateInspectionSnapshot();
            string text = before.ToText();
            string normalized = text.Replace("\r\n", "\n");

            Assert.Contains(@"Line 1\nLine 2", text, StringComparison.Ordinal);
            Assert.Contains(@"data[owner\=team\nprimary]=value=one\nvalue two", text, StringComparison.Ordinal);
            Assert.DoesNotContain("Line 1\nLine 2", normalized, StringComparison.Ordinal);

            node.Text = "Line 1\nChanged";
            VisioInspectionDiff diff = before.Diff(document.CreateInspectionSnapshot());

            Assert.Contains(diff.Differences, difference =>
                difference.Kind == VisioInspectionDifferenceKind.Changed &&
                difference.Path == "page[" + page.Id + ":Escapes].shape[" + node.Id + "].text" &&
                difference.Expected == "Line 1\nLine 2" &&
                difference.Actual == "Line 1\nChanged");
            Assert.Contains(@"expected=Line 1\nLine 2 actual=Line 1\nChanged", diff.ToText(), StringComparison.Ordinal);
        }

        [Fact]
        public void InspectionSnapshotPagePathsIncludePageIdForDuplicateNames() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage first = document.AddPage("Same", 6, 4);
            VisioPage second = document.AddPage("Same", 6, 4);
            VisioShape firstShape = first.AddRectangle(1, 2, 1, 0.5, "First");
            VisioShape secondShape = second.AddRectangle(1, 2, 1, 0.5, "Second");

            string text = document.CreateInspectionSnapshot().ToText();

            Assert.Contains("page[" + first.Id + ":Same].shape[" + firstShape.Id + "].text=First", text, StringComparison.Ordinal);
            Assert.Contains("page[" + second.Id + ":Same].shape[" + secondShape.Id + "].text=Second", text, StringComparison.Ordinal);
        }

        [Fact]
        public void InspectionDiffReportsConnectorLabelCoordinateMovement() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Labels", 6, 4);
            VisioShape start = page.AddRectangle(1, 2, 1, 0.5, "Start");
            VisioShape end = page.AddRectangle(4, 2, 1, 0.5, "End");
            VisioConnector connector = page.AddConnector(start, end, ConnectorKind.RightAngle, VisioSide.Right, VisioSide.Left)
                .PlaceLabel(0.5);
            connector.Label = "Moved";

            VisioInspectionSnapshot before = document.CreateInspectionSnapshot();

            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(0.5, offsetY: 0.25);
            VisioInspectionDiff diff = before.Diff(document.CreateInspectionSnapshot());

            Assert.Contains(diff.Differences, difference =>
                difference.Kind == VisioInspectionDifferenceKind.Changed &&
                difference.Path == "page[" + page.Id + ":Labels].connector[" + connector.Id + "].labelResolvedPinY" &&
                difference.Expected == "2" &&
                difference.Actual == "2.25");
        }
    }
}
