using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    public static partial class VisioConnectorRoutingExtensions {
        private static void ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint, out double x, out double y) {
            if (connectionPoint != null) {
                (x, y) = GetPagePoint(shape, connectionPoint.X, connectionPoint.Y);
                return;
            }

            VisioShapeBounds shapeBounds = GetPageShapeBounds(shape);
            VisioShapeBounds otherBounds = GetPageShapeBounds(other);
            double left = shapeBounds.Left;
            double bottom = shapeBounds.Bottom;
            double right = shapeBounds.Right;
            double top = shapeBounds.Top;
            double otherLeft = otherBounds.Left;
            double otherBottom = otherBounds.Bottom;
            double otherRight = otherBounds.Right;
            double otherTop = otherBounds.Top;
            double cx = (left + right) / 2D;
            double cy = (bottom + top) / 2D;
            double otherCx = (otherLeft + otherRight) / 2D;
            double otherCy = (otherBottom + otherTop) / 2D;
            double dx = otherCx - cx;
            double dy = otherCy - cy;

            if (Math.Abs(dx) >= Math.Abs(dy)) {
                x = dx >= 0 ? right : left;
                y = cy;
            } else {
                x = cx;
                y = dy >= 0 ? top : bottom;
            }
        }

        private static IReadOnlyList<VisioConnector> OrderConnectorsForPageRouting(IReadOnlyList<VisioConnector> connectors, IEnumerable<VisioShape> obstacles, VisioConnectorRoutingOptions options) {
            List<ConnectorRoutingWorkItem> workItems = new();
            int index = 0;
            foreach (VisioConnector connector in connectors) {
                ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
                ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
                IEnumerable<VisioShape> routingObstacles = options.IncludeGroupChildren
                    ? ExpandRoutingObstacles(obstacles)
                    : obstacles;
                List<VisioShapeBounds> obstacleBounds = GetRoutingObstacleBounds(connector, routingObstacles, options.Padding, options);
                List<IReadOnlyList<RoutePoint>> connectorReferencePaths = options.AvoidConnectorCrossings
                    ? GetConnectorReferencePaths(connector, options.ConnectorCrossingReferences)
                    : new List<IReadOnlyList<RoutePoint>>();
                RouteScore score = ScoreCurrentRoute(connector, startX, startY, endX, endY, obstacleBounds, connectorReferencePaths);
                workItems.Add(new ConnectorRoutingWorkItem(connector, score, index++));
            }

            return workItems
                .OrderByDescending(item => item.Score.Intersections)
                .ThenByDescending(item => item.Score.ConnectorCrossings)
                .ThenByDescending(item => item.Score.Length)
                .ThenBy(item => item.Index)
                .Select(item => item.Connector)
                .ToList();
        }

        private static RouteScore ScorePageRoutes(IReadOnlyList<VisioConnector> connectors, IEnumerable<VisioShape> obstacles, VisioConnectorRoutingOptions options) {
            int intersections = 0;
            double length = 0D;
            foreach (VisioConnector connector in connectors) {
                ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
                ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
                IEnumerable<VisioShape> routingObstacles = options.IncludeGroupChildren
                    ? ExpandRoutingObstacles(obstacles)
                    : obstacles;
                List<VisioShapeBounds> obstacleBounds = GetRoutingObstacleBounds(connector, routingObstacles, options.Padding, options);
                RouteScore score = ScoreCurrentRoute(connector, startX, startY, endX, endY, obstacleBounds, new List<IReadOnlyList<RoutePoint>>());
                intersections += score.Intersections;
                length += score.Length;
            }

            return new RouteScore(intersections, CountPageConnectorCrossings(connectors), length);
        }

        private static List<VisioShapeBounds> GetRoutingObstacleBounds(VisioConnector connector, IEnumerable<VisioShape> obstacles, double padding, VisioConnectorRoutingOptions options) {
            List<VisioShapeBounds> bounds = new();
            VisioShapeBounds fromBounds = GetPageShapeBounds(connector.From);
            VisioShapeBounds toBounds = GetPageShapeBounds(connector.To);
            foreach (VisioShape obstacle in obstacles) {
                if (IsEndpointRelated(obstacle, connector.From) ||
                    IsEndpointRelated(obstacle, connector.To)) {
                    continue;
                }

                if (obstacle.IsContainer && !options.IncludeContainers) {
                    continue;
                }

                if (obstacle.IsBackgroundSurface && !options.IncludeBackgroundSurfaces) {
                    continue;
                }

                if (VisioSemanticUserCells.IsGeneratedDiagramAdornment(obstacle) && !options.IncludeDiagramAdornments) {
                    continue;
                }

                VisioShapeBounds shapeBounds = GetPageShapeBounds(obstacle);
                if (!shapeBounds.IsEmpty &&
                    !Contains(shapeBounds, fromBounds) &&
                    !Contains(shapeBounds, toBounds)) {
                    bounds.Add(Inflate(shapeBounds, padding));
                }
            }

            return bounds;
        }

        private static IReadOnlyList<VisioShape> ExpandRoutingObstacles(IEnumerable<VisioShape> obstacles) {
            List<VisioShape> expanded = new();
            HashSet<VisioShape> seen = new();
            foreach (VisioShape obstacle in obstacles) {
                AddObstacleAndChildren(obstacle);
            }

            return expanded;

            void AddObstacleAndChildren(VisioShape obstacle) {
                if (!seen.Add(obstacle)) {
                    return;
                }

                expanded.Add(obstacle);
                foreach (VisioShape child in obstacle.Children) {
                    AddObstacleAndChildren(child);
                }
            }
        }

        private static bool IsEndpointRelated(VisioShape obstacle, VisioShape endpoint) {
            return ReferenceEquals(obstacle, endpoint) ||
                   IsAncestorOf(obstacle, endpoint) ||
                   IsAncestorOf(endpoint, obstacle);
        }

        private static bool IsAncestorOf(VisioShape possibleAncestor, VisioShape shape) {
            for (VisioShape? parent = shape.Parent; parent != null; parent = parent.Parent) {
                if (ReferenceEquals(parent, possibleAncestor)) {
                    return true;
                }
            }

            return false;
        }

        private static VisioShapeBounds GetPageShapeBounds(VisioShape shape) {
            (double x1, double y1) = GetPagePoint(shape, 0, 0);
            (double x2, double y2) = GetPagePoint(shape, shape.Width, 0);
            (double x3, double y3) = GetPagePoint(shape, 0, shape.Height);
            (double x4, double y4) = GetPagePoint(shape, shape.Width, shape.Height);
            double left = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double right = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double bottom = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double top = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));
            return new VisioShapeBounds(left, bottom, right, top);
        }

        private static (double X, double Y) GetPagePoint(VisioShape shape, double x, double y) {
            (double absX, double absY) = shape.GetAbsolutePoint(x, y);
            return shape.Parent != null
                ? GetPagePoint(shape.Parent, absX, absY)
                : (absX, absY);
        }

        private static IEnumerable<RouteCandidate> EnumerateOrthogonalRouteCandidates(double startX, double startY, double endX, double endY, double step, int maxLanes) {
            VisioConnectorRouteStyle primary = Math.Abs(endX - startX) >= Math.Abs(endY - startY)
                ? VisioConnectorRouteStyle.VerticalThenHorizontal
                : VisioConnectorRouteStyle.HorizontalThenVertical;
            VisioConnectorRouteStyle secondary = primary == VisioConnectorRouteStyle.VerticalThenHorizontal
                ? VisioConnectorRouteStyle.HorizontalThenVertical
                : VisioConnectorRouteStyle.VerticalThenHorizontal;

            foreach (double offset in EnumerateLaneOffsets(step, maxLanes)) {
                yield return CreateOrthogonalRouteCandidate(startX, startY, endX, endY, primary, offset);
                yield return CreateOrthogonalRouteCandidate(startX, startY, endX, endY, secondary, offset);
            }

            double[] offsets = EnumerateLaneOffsets(step, maxLanes).ToArray();
            foreach (double xOffset in offsets) {
                foreach (double yOffset in offsets) {
                    if (Math.Abs(xOffset) < 1e-9 && Math.Abs(yOffset) < 1e-9) {
                        continue;
                    }

                    yield return CreateDoglegRouteCandidate(startX, startY, endX, endY, xOffset, yOffset, true);
                    yield return CreateDoglegRouteCandidate(startX, startY, endX, endY, xOffset, yOffset, false);
                }
            }
        }

        private static IEnumerable<double> EnumerateLaneOffsets(double step, int maxLanes) {
            double resolvedStep = step > 0D ? step : 0.15D;
            yield return 0D;
            for (int lane = 1; lane <= maxLanes; lane++) {
                double offset = lane * resolvedStep;
                yield return offset;
                yield return -offset;
            }
        }

        private static RouteCandidate CreateOrthogonalRouteCandidate(double startX, double startY, double endX, double endY, VisioConnectorRouteStyle style, double offset) {
            VisioConnectorRouteStyle resolvedStyle = style == VisioConnectorRouteStyle.Auto
                ? Math.Abs(endX - startX) >= Math.Abs(endY - startY)
                    ? VisioConnectorRouteStyle.HorizontalThenVertical
                    : VisioConnectorRouteStyle.VerticalThenHorizontal
                : style;

            if (resolvedStyle == VisioConnectorRouteStyle.HorizontalThenVertical) {
                double laneX = ((startX + endX) / 2D) + offset;
                return new RouteCandidate(
                    new RoutePoint(startX, startY),
                    new RoutePoint(laneX, startY),
                    new RoutePoint(laneX, endY),
                    new RoutePoint(endX, endY));
            }

            double laneY = ((startY + endY) / 2D) + offset;
            return new RouteCandidate(
                new RoutePoint(startX, startY),
                new RoutePoint(startX, laneY),
                new RoutePoint(endX, laneY),
                new RoutePoint(endX, endY));
        }

        private static RouteCandidate CreateDoglegRouteCandidate(double startX, double startY, double endX, double endY, double xOffset, double yOffset, bool horizontalEscapeFirst) {
            double laneX = ((startX + endX) / 2D) + xOffset;
            double laneY = ((startY + endY) / 2D) + yOffset;
            if (horizontalEscapeFirst) {
                return new RouteCandidate(
                    new RoutePoint(startX, startY),
                    new RoutePoint(laneX, startY),
                    new RoutePoint(laneX, laneY),
                    new RoutePoint(endX, laneY),
                    new RoutePoint(endX, endY));
            }

            return new RouteCandidate(
                new RoutePoint(startX, startY),
                new RoutePoint(startX, laneY),
                new RoutePoint(laneX, laneY),
                new RoutePoint(laneX, endY),
                new RoutePoint(endX, endY));
        }

        private static RouteScore ScoreRoute(RouteCandidate candidate, IReadOnlyList<VisioShapeBounds> obstacles, IReadOnlyList<IReadOnlyList<RoutePoint>> connectorReferencePaths) {
            int intersections = 0;
            foreach (VisioShapeBounds obstacle in obstacles) {
                if (RouteIntersectsBounds(candidate, obstacle)) {
                    intersections++;
                }
            }

            return new RouteScore(intersections, CountConnectorCrossings(candidate, connectorReferencePaths), candidate.Length);
        }

        private static RouteScore ScoreCurrentRoute(VisioConnector connector, double startX, double startY, double endX, double endY, IReadOnlyList<VisioShapeBounds> obstacles, IReadOnlyList<IReadOnlyList<RoutePoint>> connectorReferencePaths) {
            List<RoutePoint> points = GetConnectorPath(connector, startX, startY, endX, endY);
            int intersections = 0;
            foreach (VisioShapeBounds obstacle in obstacles) {
                if (PathIntersectsBounds(points, obstacle)) {
                    intersections++;
                }
            }

            double length = 0D;
            for (int i = 1; i < points.Count; i++) {
                length += Distance(points[i - 1], points[i]);
            }

            return new RouteScore(intersections, CountConnectorCrossings(points, connectorReferencePaths), length);
        }

        private static List<RoutePoint> GetConnectorPath(VisioConnector connector, double startX, double startY, double endX, double endY) {
            List<(double X, double Y)> waypoints = connector.Waypoints
                .Select(waypoint => (X: waypoint.X, Y: waypoint.Y))
                .ToList();

            return OfficeGeometry.BuildConnectorPolyline(
                    (startX, startY),
                    (endX, endY),
                    waypoints,
                    connector.Kind == ConnectorKind.RightAngle)
                .Select(point => new RoutePoint(point.X, point.Y))
                .ToList();
        }

        private static List<IReadOnlyList<RoutePoint>> GetConnectorReferencePaths(VisioConnector connector, IEnumerable<VisioConnector>? referenceConnectors) {
            List<IReadOnlyList<RoutePoint>> paths = new();
            if (referenceConnectors == null) {
                return paths;
            }

            foreach (VisioConnector reference in referenceConnectors) {
                if (reference == null || ReferenceEquals(reference, connector)) {
                    continue;
                }

                ResolveEndpoint(reference.From, reference.To, reference.FromConnectionPoint, out double startX, out double startY);
                ResolveEndpoint(reference.To, reference.From, reference.ToConnectionPoint, out double endX, out double endY);
                List<RoutePoint> path = GetConnectorPath(reference, startX, startY, endX, endY);
                if (path.Count > 1) {
                    paths.Add(path);
                }
            }

            return paths;
        }

        private static bool RouteIntersectsBounds(RouteCandidate route, VisioShapeBounds bounds) {
            return PathIntersectsBounds(route.Points, bounds);
        }

        private static bool PathIntersectsBounds(IReadOnlyList<RoutePoint> points, VisioShapeBounds bounds) {
            for (int i = 1; i < points.Count; i++) {
                if (SegmentIntersectsBounds(points[i - 1], points[i], bounds)) {
                    return true;
                }
            }

            return false;
        }

        private static bool SegmentIntersectsBounds(RoutePoint a, RoutePoint b, VisioShapeBounds bounds) {
            return OfficeGeometry.SegmentIntersectsRectangle(
                (a.X, a.Y),
                (b.X, b.Y),
                bounds.Left,
                bounds.Bottom,
                bounds.Right,
                bounds.Top);
        }

        private static int CountConnectorCrossings(RouteCandidate candidate, IReadOnlyList<IReadOnlyList<RoutePoint>> connectorReferencePaths) {
            return CountConnectorCrossings(candidate.Points, connectorReferencePaths);
        }

        private static int CountPageConnectorCrossings(IReadOnlyList<VisioConnector> connectors) {
            List<IReadOnlyList<RoutePoint>> paths = new();
            foreach (VisioConnector connector in connectors) {
                ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
                ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
                paths.Add(GetConnectorPath(connector, startX, startY, endX, endY));
            }

            int crossings = 0;
            for (int i = 0; i < paths.Count; i++) {
                for (int j = i + 1; j < paths.Count; j++) {
                    crossings += CountConnectorCrossings(paths[i], new[] { paths[j] });
                }
            }

            return crossings;
        }

        private static int CountConnectorCrossings(IReadOnlyList<RoutePoint> points, IReadOnlyList<IReadOnlyList<RoutePoint>> connectorReferencePaths) {
            if (connectorReferencePaths.Count == 0) {
                return 0;
            }

            int crossings = 0;
            for (int i = 1; i < points.Count; i++) {
                RoutePoint from = points[i - 1];
                RoutePoint to = points[i];
                foreach (IReadOnlyList<RoutePoint> referencePath in connectorReferencePaths) {
                    for (int j = 1; j < referencePath.Count; j++) {
                        if (SegmentsIntersectAwayFromSharedEndpoints(from, to, referencePath[j - 1], referencePath[j])) {
                            crossings++;
                        }
                    }
                }
            }

            return crossings;
        }

        private static bool SegmentsIntersectAwayFromSharedEndpoints(RoutePoint p1, RoutePoint p2, RoutePoint q1, RoutePoint q2) {
            return OfficeGeometry.SegmentsIntersect((p1.X, p1.Y), (p2.X, p2.Y), (q1.X, q1.Y), (q2.X, q2.Y)) &&
                   !PointsEqual(p1, q1) &&
                   !PointsEqual(p1, q2) &&
                   !PointsEqual(p2, q1) &&
                   !PointsEqual(p2, q2);
        }

        private static bool PointsEqual(RoutePoint a, RoutePoint b) {
            return Math.Abs(a.X - b.X) < 1e-9 &&
                   Math.Abs(a.Y - b.Y) < 1e-9;
        }

        private static bool Contains(VisioShapeBounds outer, VisioShapeBounds inner) {
            if (outer.IsEmpty || inner.IsEmpty) {
                return false;
            }

            const double tolerance = 1e-6;
            return outer.Left <= inner.Left + tolerance &&
                   outer.Bottom <= inner.Bottom + tolerance &&
                   outer.Right + tolerance >= inner.Right &&
                   outer.Top + tolerance >= inner.Top;
        }

        private static VisioShapeBounds Inflate(VisioShapeBounds bounds, double padding) {
            return new VisioShapeBounds(
                bounds.Left - padding,
                bounds.Bottom - padding,
                bounds.Right + padding,
                bounds.Top + padding);
        }
    }
}
