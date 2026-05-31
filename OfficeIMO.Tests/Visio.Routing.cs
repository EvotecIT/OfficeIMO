using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Fluent;
using OfficeIMO.Visio.Stencils;
using Xunit;
using VisioBounds = OfficeIMO.Visio.VisioShapeBounds;

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
        public void ObstacleAwareRoutingCanAvoidUnrelatedBackgroundZones() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Zone routes", 8, 5);
            VisioShape source = page.AddRectangle(1, 2.5, 0.8, 0.5, "Source");
            VisioShape zone = page.AddRectangle(4, 2.5, 2.0, 1.0, "Restricted zone");
            zone.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            VisioShape target = page.AddRectangle(7, 2.5, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            connector.RouteOrthogonalAroundShapes(page.Shapes);

            Assert.Empty(connector.Waypoints);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);

            connector.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions {
                IncludeBackgroundSurfaces = true,
                Padding = 0.15D,
                MaxLanes = 16
            });

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            VisioBounds paddedZone = Inflate(zone.GetShapeBounds(), 0.15D);
            Assert.DoesNotContain(connector.Waypoints, waypoint =>
                waypoint.X > paddedZone.Left &&
                waypoint.X < paddedZone.Right &&
                waypoint.Y > paddedZone.Bottom &&
                waypoint.Y < paddedZone.Top);
            Assert.DoesNotContain(GetRouteSegments(connector), segment => SegmentIntersectsBounds(segment.Start, segment.End, paddedZone));
        }

        [Fact]
        public void PolishDiagramCanRouteConnectorsAroundBackgroundZonesWhenRequested() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Zone polish", 8, 5);
            VisioShape source = page.AddRectangle(1, 2.5, 0.8, 0.5, "Source");
            VisioShape zone = page.AddRectangle(4, 2.5, 2.0, 1.0, "Restricted zone");
            zone.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.BackgroundSurfaceKind, "STR");
            VisioShape target = page.AddRectangle(7, 2.5, 0.8, 0.5, "Target");
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ResolveConnectorShapeIntersections = true,
                ConnectorRoutingAvoidBackgroundSurfaces = true,
                ConnectorRoutingMaxLanes = 16,
                FitToContent = false
            });

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            Assert.DoesNotContain(GetRouteSegments(connector), segment => SegmentIntersectsBounds(segment.Start, segment.End, Inflate(zone.GetShapeBounds(), 0.15D)));
        }

        [Fact]
        public void ObstacleAwareRoutingCanAvoidSiblingShapesInsideEndpointGroup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Group routes", 9, 5);
            VisioShape source = page.AddRectangle(1.2, 2.5, 0.8, 0.5, "Source");
            VisioShape group = page.AddRectangle(5.1, 2.5, 4.4, 2.2, "Service group");
            VisioShape blocker = new("member-blocker", 4.2, 2.5, 1.0, 0.9, "Blocker");
            VisioShape target = new("member-target", 7.1, 2.5, 0.8, 0.5, "Target");
            group.Children.Add(blocker);
            group.Children.Add(target);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            connector.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions {
                Padding = 0.12D,
                MaxLanes = 16
            });

            Assert.Empty(connector.Waypoints);
            Assert.Equal(ConnectorKind.Straight, connector.Kind);

            connector.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions {
                IncludeGroupChildren = true,
                Padding = 0.12D,
                MaxLanes = 16
            });

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            Assert.DoesNotContain(GetRouteSegments(connector), segment => SegmentIntersectsBounds(segment.Start, segment.End, Inflate(blocker.GetShapeBounds(), 0.12D)));

            document.Save();
            Assert.Empty(VisioValidator.Validate(filePath));
        }

        [Fact]
        public void PolishDiagramCanRouteAroundGroupChildrenWhenRequested() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Group polish", 9, 5);
            VisioShape source = page.AddRectangle(1.2, 2.5, 0.8, 0.5, "Source");
            VisioShape group = page.AddRectangle(5.1, 2.5, 4.4, 2.2, "Service group");
            VisioShape blocker = new("member-blocker", 4.2, 2.5, 1.0, 0.9, "Blocker");
            VisioShape target = new("member-target", 7.1, 2.5, 0.8, 0.5, "Target");
            group.Children.Add(blocker);
            group.Children.Add(target);
            VisioConnector connector = page.AddConnector(source, target, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ResolveConnectorShapeIntersections = true,
                ConnectorRoutingAvoidGroupChildren = true,
                ConnectorRoutingMaxLanes = 16,
                FitToContent = false
            });

            Assert.Equal(ConnectorKind.RightAngle, connector.Kind);
            Assert.Equal(2, connector.Waypoints.Count);
            Assert.DoesNotContain(GetRouteSegments(connector), segment => SegmentIntersectsBounds(segment.Start, segment.End, Inflate(blocker.GetShapeBounds(), 0.15D)));
        }

        [Fact]
        public void ObstacleAwareRoutingCanReduceConnectorCrossingsWhenRequested() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Crossing routes", 8, 6);
            VisioShape left = page.AddRectangle(1, 3, 0.8, 0.5, "Left");
            VisioShape right = page.AddRectangle(7, 3, 0.8, 0.5, "Right");
            VisioShape top = page.AddRectangle(4, 5, 0.8, 0.5, "Top");
            VisioShape bottom = page.AddRectangle(4, 1, 0.8, 0.5, "Bottom");
            VisioConnector horizontal = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            VisioConnector vertical = page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            Assert.True(CountRouteCrossings(vertical, horizontal) > 0);

            vertical.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions {
                AvoidConnectorCrossings = true,
                ConnectorCrossingReferences = page.Connectors,
                Padding = 0.15D,
                MaxLanes = 24
            });

            Assert.Equal(ConnectorKind.RightAngle, vertical.Kind);
            Assert.Equal(2, vertical.Waypoints.Count);
            Assert.Equal(0, CountRouteCrossings(vertical, horizontal));
        }

        [Fact]
        public void PolishDiagramCanReduceConnectorCrossingsWhenRequested() {
            VisioDocument document = VisioDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx"));
            VisioPage page = document.AddPage("Crossing polish", 8, 6);
            VisioShape left = page.AddRectangle(1, 3, 0.8, 0.5, "Left");
            VisioShape right = page.AddRectangle(7, 3, 0.8, 0.5, "Right");
            VisioShape top = page.AddRectangle(4, 5, 0.8, 0.5, "Top");
            VisioShape bottom = page.AddRectangle(4, 1, 0.8, 0.5, "Bottom");
            VisioConnector horizontal = page.AddConnector(left, right, ConnectorKind.Straight, VisioSide.Right, VisioSide.Left);
            VisioConnector vertical = page.AddConnector(top, bottom, ConnectorKind.Straight, VisioSide.Bottom, VisioSide.Top);

            Assert.True(CountRouteCrossings(vertical, horizontal) > 0);

            page.PolishDiagram(new VisioDiagramPolishOptions {
                ResolveConnectorShapeIntersections = true,
                ConnectorRoutingAvoidConnectorCrossings = true,
                ConnectorRoutingMaxLanes = 24,
                FitToContent = false
            });

            Assert.Equal(0, CountRouteCrossings(vertical, horizontal));
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

        private static RouteSegment[] GetRouteSegments(VisioConnector connector) {
            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out RoutePoint start);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out RoutePoint end);
            RoutePoint[] points = new[] { start }
                .Concat(connector.Waypoints.Select(waypoint => new RoutePoint(waypoint.X, waypoint.Y)))
                .Concat(new[] { end })
                .ToArray();
            return points.Zip(points.Skip(1), (from, to) => new RouteSegment(from, to)).ToArray();
        }

        private static void ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint, out RoutePoint point) {
            if (connectionPoint != null) {
                (double x, double y) = shape.GetAbsolutePoint(connectionPoint.X, connectionPoint.Y);
                point = new RoutePoint(x, y);
                return;
            }

            VisioBounds shapeBounds = shape.GetShapeBounds();
            VisioBounds otherBounds = other.GetShapeBounds();
            double dx = otherBounds.CenterX - shapeBounds.CenterX;
            double dy = otherBounds.CenterY - shapeBounds.CenterY;
            if (Math.Abs(dx) >= Math.Abs(dy)) {
                point = new RoutePoint(dx >= 0 ? shapeBounds.Right : shapeBounds.Left, shapeBounds.CenterY);
            } else {
                point = new RoutePoint(shapeBounds.CenterX, dy >= 0 ? shapeBounds.Top : shapeBounds.Bottom);
            }
        }

        private static bool SegmentIntersectsBounds(RoutePoint a, RoutePoint b, VisioBounds bounds) {
            if (PointInside(a, bounds) || PointInside(b, bounds)) {
                return true;
            }

            RoutePoint bottomLeft = new(bounds.Left, bounds.Bottom);
            RoutePoint bottomRight = new(bounds.Right, bounds.Bottom);
            RoutePoint topLeft = new(bounds.Left, bounds.Top);
            RoutePoint topRight = new(bounds.Right, bounds.Top);

            return SegmentsIntersect(a, b, bottomLeft, bottomRight) ||
                   SegmentsIntersect(a, b, bottomRight, topRight) ||
                   SegmentsIntersect(a, b, topRight, topLeft) ||
                   SegmentsIntersect(a, b, topLeft, bottomLeft);
        }

        private static int CountRouteCrossings(VisioConnector connector, params VisioConnector[] references) {
            RouteSegment[] segments = GetRouteSegments(connector);
            int crossings = 0;
            foreach (RouteSegment segment in segments) {
                foreach (VisioConnector reference in references) {
                    foreach (RouteSegment referenceSegment in GetRouteSegments(reference)) {
                        if (SegmentsIntersectAwayFromSharedEndpoints(segment, referenceSegment)) {
                            crossings++;
                        }
                    }
                }
            }

            return crossings;
        }

        private static bool SegmentsIntersectAwayFromSharedEndpoints(RouteSegment segment, RouteSegment referenceSegment) {
            return SegmentsIntersect(segment.Start, segment.End, referenceSegment.Start, referenceSegment.End) &&
                   !PointsEqual(segment.Start, referenceSegment.Start) &&
                   !PointsEqual(segment.Start, referenceSegment.End) &&
                   !PointsEqual(segment.End, referenceSegment.Start) &&
                   !PointsEqual(segment.End, referenceSegment.End);
        }

        private static bool SegmentsIntersect(RoutePoint p1, RoutePoint p2, RoutePoint q1, RoutePoint q2) {
            double o1 = Orientation(p1, p2, q1);
            double o2 = Orientation(p1, p2, q2);
            double o3 = Orientation(q1, q2, p1);
            double o4 = Orientation(q1, q2, p2);

            if (o1 * o2 < 0D && o3 * o4 < 0D) {
                return true;
            }

            return IsZero(o1) && OnSegment(p1, q1, p2) ||
                   IsZero(o2) && OnSegment(p1, q2, p2) ||
                   IsZero(o3) && OnSegment(q1, p1, q2) ||
                   IsZero(o4) && OnSegment(q1, p2, q2);
        }

        private static double Orientation(RoutePoint a, RoutePoint b, RoutePoint c) {
            return ((b.X - a.X) * (c.Y - a.Y)) - ((b.Y - a.Y) * (c.X - a.X));
        }

        private static bool OnSegment(RoutePoint a, RoutePoint b, RoutePoint c) {
            return b.X >= Math.Min(a.X, c.X) - 1e-9 &&
                   b.X <= Math.Max(a.X, c.X) + 1e-9 &&
                   b.Y >= Math.Min(a.Y, c.Y) - 1e-9 &&
                   b.Y <= Math.Max(a.Y, c.Y) + 1e-9;
        }

        private static bool PointsEqual(RoutePoint a, RoutePoint b) {
            return Math.Abs(a.X - b.X) < 1e-9 &&
                   Math.Abs(a.Y - b.Y) < 1e-9;
        }

        private static bool PointInside(RoutePoint point, VisioBounds bounds) {
            return point.X > bounds.Left && point.X < bounds.Right &&
                   point.Y > bounds.Bottom && point.Y < bounds.Top;
        }

        private static bool IsZero(double value) {
            return Math.Abs(value) < 1e-9;
        }

        private static VisioBounds Inflate(VisioBounds bounds, double padding) {
            return new VisioBounds(
                bounds.Left - padding,
                bounds.Bottom - padding,
                bounds.Right + padding,
                bounds.Top + padding);
        }

        private readonly struct RoutePoint {
            public RoutePoint(double x, double y) {
                X = x;
                Y = y;
            }

            public double X { get; }

            public double Y { get; }
        }

        private readonly struct RouteSegment {
            public RouteSegment(RoutePoint start, RoutePoint end) {
                Start = start;
                End = end;
            }

            public RoutePoint Start { get; }

            public RoutePoint End { get; }
        }
    }
}
