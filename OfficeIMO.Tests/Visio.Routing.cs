using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioRoutingTests {
        [Fact]
        public void ExplicitConnectorWaypointsAreWrittenToVsdxGeometry() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Routes");
            VisioShape source = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "source", 2, 5, "Source");
            VisioShape target = page.AddStencilShape(VisioStencils.BasicShapes.Get("rectangle"), "target", 7, 3, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteThrough(
                    VisioConnectorWaypoint.At(4.5, 5),
                    VisioConnectorWaypoint.At(4.5, 3));
            connector.EndArrow = EndArrow.Triangle;
            connector.Label = "manual route";

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));

            XDocument pageXml = ReadPageXml(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = pageXml
                .Descendants(ns + "Shape")
                .Single(shape =>
                    string.Equals((string?)shape.Attribute("NameU"), "Connector", StringComparison.Ordinal) ||
                    string.Equals((string?)shape.Attribute("NameU"), "Dynamic connector", StringComparison.Ordinal));
            XElement geometry = connectorShape
                .Elements(ns + "Section")
                .Single(section => string.Equals((string?)section.Attribute("N"), "Geometry", StringComparison.Ordinal));
            var lineToRows = geometry
                .Elements(ns + "Row")
                .Where(row => string.Equals((string?)row.Attribute("T"), "LineTo", StringComparison.Ordinal))
                .ToArray();

            Assert.Equal(3, lineToRows.Length);
            Assert.Equal("4.5", CellValue(lineToRows[0], ns, "X"));
            Assert.Equal("5", CellValue(lineToRows[0], ns, "Y"));
            Assert.Equal("4.5", CellValue(lineToRows[1], ns, "X"));
            Assert.Equal("3", CellValue(lineToRows[1], ns, "Y"));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            Assert.Equal(ConnectorKind.RightAngle, loadedConnector.Kind);
            Assert.Equal(2, loadedConnector.Waypoints.Count);
            Assert.Equal(4.5, loadedConnector.Waypoints[0].X, 6);
            Assert.Equal(5, loadedConnector.Waypoints[0].Y, 6);
            Assert.Equal(4.5, loadedConnector.Waypoints[1].X, 6);
            Assert.Equal(3, loadedConnector.Waypoints[1].Y, 6);

            loaded.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
            VisioConnector roundTrippedConnector = VisioDocument.Load(filePath).Pages[0].Connectors.Single();
            Assert.Equal(2, roundTrippedConnector.Waypoints.Count);
        }

        [Fact]
        public void OrthogonalRoutingWorksFromModelSelectionAndFluentApi() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Route editing");
            VisioShape intake = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "intake", 2, 5, "Intake");
            VisioShape review = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "review", 6, 5, "Review");
            VisioShape archive = page.AddStencilShape(VisioStencils.Flowchart.Get("data"), "archive", 6, 2, "Archive");
            VisioConnector first = page.AddConnector(intake, review, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left);
            VisioConnector second = page.AddConnector(review, archive, ConnectorKind.Dynamic, VisioSide.Bottom, VisioSide.Top);

            first.RouteOrthogonal(VisioConnectorRouteStyle.HorizontalThenVertical, 0.25);
            page.SelectOutgoingConnectors(review)
                .RouteOrthogonal(VisioConnectorRouteStyle.VerticalThenHorizontal, -0.2)
                .EndArrow(EndArrow.Triangle);

            Assert.Equal(2, first.Waypoints.Count);
            Assert.Equal(2, second.Waypoints.Count);
            Assert.Equal(ConnectorKind.RightAngle, first.Kind);
            Assert.Equal(ConnectorKind.RightAngle, second.Kind);
            Assert.Equal(EndArrow.Triangle, second.EndArrow);

            first.ClearRoute();

            Assert.Empty(first.Waypoints);
            Assert.Equal(ConnectorKind.Dynamic, first.Kind);

            VisioDocument fluentDocument = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            fluentDocument.AsFluent()
                .Page("Fluent routes", pageBuilder => pageBuilder
                    .Rect("a", 1, 1, 1.5, 0.8, "A")
                    .Rect("b", 5, 4, 1.5, 0.8, "B")
                    .Connect("a", "b", VisioSide.Right, VisioSide.Left, connector => connector
                        .RouteThrough(VisioConnectorWaypoint.At(3, 1), VisioConnectorWaypoint.At(3, 4))
                        .ArrowEnd(EndArrow.Triangle)
                        .Label("routed")))
                .End();

            VisioConnector fluentConnector = fluentDocument.Pages[0].Connectors.Single();
            Assert.Equal(2, fluentConnector.Waypoints.Count);
            Assert.Equal("routed", fluentConnector.Label);
        }

        [Fact]
        public void ObstacleAwareOrthogonalRoutingAvoidsUnrelatedShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Obstacle routes", 8, 5);
            VisioShape source = page.AddRectangle(1, 2.5, 0.8, 0.5, "Source");
            VisioShape obstacle = page.AddRectangle(4, 2.5, 1.2, 1.0, "Obstacle");
            VisioShape target = page.AddRectangle(7, 2.5, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            Assert.Contains(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ConnectorCrossesShape" && issue.ShapeId == obstacle.Id && issue.ConnectorId == connector.Id);

            page.RouteConnectorsOrthogonalAroundShapes(padding: 0.12D, maxLanes: 16);

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            Assert.DoesNotContain(page.AnalyzeVisualQuality(new VisioDiagramQualityOptions {
                CheckShapeOverlaps = false,
                CheckConnectorLabels = false
            }), issue => issue.Kind == "ConnectorCrossesShape" && issue.ShapeId == obstacle.Id && issue.ConnectorId == connector.Id);

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void ConnectorLabelPlacementIsWrittenLoadedAndAvailableThroughSelectionsAndFluentApi() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Labels");
            VisioShape source = page.AddStencilShape(VisioStencils.Flowchart.Get("process"), "source", 2, 5, "Source");
            VisioShape decision = page.AddStencilShape(VisioStencils.Flowchart.Get("decision"), "decision", 6, 5, "Decision");
            VisioConnector connector = page.AddConnector(source, decision, ConnectorKind.Dynamic, VisioSide.Right, VisioSide.Left)
                .RouteOrthogonal()
                .PlaceLabelAt(4.25, 5.35, width: 1.4, height: 0.35);
            connector.Label = "approved";

            page.SelectConnectedConnectors(source)
                .LabelPosition(0.6, offsetY: 0.2, width: 1.5, height: 0.4);
            connector.PlaceLabelAt(4.25, 5.35, width: 1.4, height: 0.35);

            document.Save();

            Assert.Empty(VisioValidator.Validate(filePath));

            XDocument pageXml = ReadPageXml(filePath);
            XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main";
            XElement connectorShape = pageXml
                .Descendants(ns + "Shape")
                .Single(shape =>
                    string.Equals((string?)shape.Attribute("NameU"), "Connector", StringComparison.Ordinal) ||
                    string.Equals((string?)shape.Attribute("NameU"), "Dynamic connector", StringComparison.Ordinal));

            Assert.Equal("4.25", ShapeCellValue(connectorShape, ns, "TxtPinX"));
            Assert.Equal("5.35", ShapeCellValue(connectorShape, ns, "TxtPinY"));
            Assert.Equal("1.4", ShapeCellValue(connectorShape, ns, "TxtWidth"));
            Assert.Equal("0.35", ShapeCellValue(connectorShape, ns, "TxtHeight"));
            Assert.Equal("0.7", ShapeCellValue(connectorShape, ns, "TxtLocPinX"));
            Assert.Equal("0.175", ShapeCellValue(connectorShape, ns, "TxtLocPinY"));

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioConnector loadedConnector = loaded.Pages[0].Connectors.Single();
            Assert.Equal("approved", loadedConnector.Label);
            Assert.NotNull(loadedConnector.LabelPlacement);
            Assert.Equal(4.25, loadedConnector.LabelPlacement!.AbsolutePinX);
            Assert.Equal(5.35, loadedConnector.LabelPlacement.AbsolutePinY);
            Assert.Equal(1.4, loadedConnector.LabelPlacement.Width);
            Assert.Equal(0.35, loadedConnector.LabelPlacement.Height);

            VisioDocument fluentDocument = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            fluentDocument.AsFluent()
                .Page("Fluent labels", pageBuilder => pageBuilder
                    .Rect("a", 1, 1, 1.5, 0.8, "A")
                    .Rect("b", 5, 1, 1.5, 0.8, "B")
                    .Connect("a", "b", VisioSide.Right, VisioSide.Left, route => route
                        .RouteOrthogonal()
                        .Label("yes", 0.7, offsetY: 0.15)))
                .End();

            VisioConnector fluentConnector = fluentDocument.Pages[0].Connectors.Single();
            Assert.Equal("yes", fluentConnector.Label);
            Assert.NotNull(fluentConnector.LabelPlacement);
            Assert.Equal(0.7, fluentConnector.LabelPlacement!.Position);
            Assert.Equal(0.15, fluentConnector.LabelPlacement.OffsetY);
        }

        private static XDocument ReadPageXml(string filePath) {
            using ZipArchive archive = ZipFile.OpenRead(filePath);
            ZipArchiveEntry? entry = archive.GetEntry("visio/pages/page1.xml");
            Assert.NotNull(entry);
            using Stream stream = entry!.Open();
            return XDocument.Load(stream);
        }

        private static string? CellValue(XElement row, XNamespace ns, string name) {
            return row
                .Elements(ns + "Cell")
                .Single(cell => string.Equals((string?)cell.Attribute("N"), name, StringComparison.Ordinal))
                .Attribute("V")
                ?.Value;
        }

        private static string? ShapeCellValue(XElement shape, XNamespace ns, string name) {
            return shape
                .Elements(ns + "Cell")
                .Single(cell => string.Equals((string?)cell.Attribute("N"), name, StringComparison.Ordinal))
                .Attribute("V")
                ?.Value;
        }
    }
}
