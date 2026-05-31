using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helpers for deterministic connector routing.
    /// </summary>
    public static class VisioConnectorRoutingExtensions {
        /// <summary>
        /// Replaces connector geometry with explicit page-coordinate waypoints.
        /// </summary>
        /// <param name="connector">Connector to route.</param>
        /// <param name="waypoints">Absolute page coordinates between start and end.</param>
        public static VisioConnector RouteThrough(this VisioConnector connector, params VisioConnectorWaypoint[] waypoints) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            return connector.RouteThrough((IEnumerable<VisioConnectorWaypoint>)waypoints);
        }

        /// <summary>
        /// Replaces connector geometry with explicit page-coordinate waypoints.
        /// </summary>
        /// <param name="connector">Connector to route.</param>
        /// <param name="waypoints">Absolute page coordinates between start and end.</param>
        public static VisioConnector RouteThrough(this VisioConnector connector, IEnumerable<VisioConnectorWaypoint> waypoints) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (waypoints == null) {
                throw new ArgumentNullException(nameof(waypoints));
            }

            connector.Waypoints.Clear();
            foreach (VisioConnectorWaypoint waypoint in waypoints) {
                if (waypoint == null) {
                    throw new ArgumentException("Route waypoints cannot contain null entries.", nameof(waypoints));
                }

                connector.Waypoints.Add(new VisioConnectorWaypoint(waypoint.X, waypoint.Y));
            }

            connector.Kind = ConnectorKind.RightAngle;
            connector.PreservedGeometrySections.Clear();
            return connector;
        }

        /// <summary>
        /// Generates a clean three-segment orthogonal route between connector endpoints.
        /// </summary>
        /// <param name="connector">Connector to route.</param>
        /// <param name="style">Orthogonal route orientation.</param>
        /// <param name="offset">Optional offset applied to the center routing lane.</param>
        public static VisioConnector RouteOrthogonal(this VisioConnector connector, VisioConnectorRouteStyle style = VisioConnectorRouteStyle.Auto, double offset = 0D) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);

            VisioConnectorRouteStyle resolvedStyle = style == VisioConnectorRouteStyle.Auto
                ? Math.Abs(endX - startX) >= Math.Abs(endY - startY)
                    ? VisioConnectorRouteStyle.HorizontalThenVertical
                    : VisioConnectorRouteStyle.VerticalThenHorizontal
                : style;

            if (resolvedStyle == VisioConnectorRouteStyle.HorizontalThenVertical) {
                double laneX = ((startX + endX) / 2D) + offset;
                return connector.RouteThrough(
                    new VisioConnectorWaypoint(laneX, startY),
                    new VisioConnectorWaypoint(laneX, endY));
            }

            double laneY = ((startY + endY) / 2D) + offset;
            return connector.RouteThrough(
                new VisioConnectorWaypoint(startX, laneY),
                new VisioConnectorWaypoint(endX, laneY));
        }

        /// <summary>
        /// Generates an orthogonal route that avoids unrelated obstacle shapes when a clear lane is available.
        /// </summary>
        /// <param name="connector">Connector to route.</param>
        /// <param name="obstacles">Shapes that the route should avoid. Source, target, containers, background surfaces, and generated adornments are ignored.</param>
        /// <param name="padding">Padding added around each obstacle while testing route intersections.</param>
        /// <param name="maxLanes">Number of positive and negative routing lanes to try on each axis.</param>
        public static VisioConnector RouteOrthogonalAroundShapes(this VisioConnector connector, IEnumerable<VisioShape> obstacles, double padding = 0.15D, int maxLanes = 12) {
            return connector.RouteOrthogonalAroundShapes(obstacles, new VisioConnectorRoutingOptions {
                Padding = padding,
                MaxLanes = maxLanes
            });
        }

        /// <summary>
        /// Generates an orthogonal route that avoids unrelated obstacle shapes when a clear lane is available.
        /// </summary>
        /// <param name="connector">Connector to route.</param>
        /// <param name="obstacles">Shapes that the route should avoid. Source and target shapes are ignored.</param>
        /// <param name="options">Routing options controlling padding, lane search, and whether zones/containers count as obstacles.</param>
        public static VisioConnector RouteOrthogonalAroundShapes(this VisioConnector connector, IEnumerable<VisioShape> obstacles, VisioConnectorRoutingOptions options) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            if (obstacles == null) {
                throw new ArgumentNullException(nameof(obstacles));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            double padding = options.Padding;
            int maxLanes = options.MaxLanes;
            if (padding < 0D || double.IsNaN(padding) || double.IsInfinity(padding)) {
                throw new ArgumentOutOfRangeException(nameof(options), "Padding must be a non-negative finite value.");
            }

            if (maxLanes < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Lane count cannot be negative.");
            }

            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
            List<VisioShapeBounds> obstacleBounds = GetRoutingObstacleBounds(connector, obstacles, padding, options);
            List<IReadOnlyList<RoutePoint>> connectorReferencePaths = options.AvoidConnectorCrossings
                ? GetConnectorReferencePaths(connector, options.ConnectorCrossingReferences)
                : new List<IReadOnlyList<RoutePoint>>();
            if (obstacleBounds.Count == 0 && connectorReferencePaths.Count == 0) {
                return connector;
            }

            RouteScore currentScore = ScoreCurrentRoute(connector, startX, startY, endX, endY, obstacleBounds, connectorReferencePaths);
            if (!currentScore.HasConflicts) {
                return connector;
            }

            RouteCandidate? best = null;
            foreach (RouteCandidate candidate in EnumerateOrthogonalRouteCandidates(startX, startY, endX, endY, padding, maxLanes)) {
                RouteScore score = ScoreRoute(candidate, obstacleBounds, connectorReferencePaths);
                if (best == null || score.IsBetterThan(best.Value.Score)) {
                    best = candidate.WithScore(score);
                }

                if (!score.HasConflicts) {
                    break;
                }
            }

            RouteCandidate resolved = best ?? CreateOrthogonalRouteCandidate(startX, startY, endX, endY, VisioConnectorRouteStyle.Auto, 0D);
            if (!resolved.Score.IsBetterThan(currentScore)) {
                return connector;
            }

            return connector.RouteThrough(
                new VisioConnectorWaypoint(resolved.First.X, resolved.First.Y),
                new VisioConnectorWaypoint(resolved.Second.X, resolved.Second.Y));
        }

        /// <summary>
        /// Routes every connector on the page around unrelated top-level shapes using deterministic orthogonal lanes.
        /// </summary>
        /// <param name="page">Page whose connectors should be rerouted.</param>
        /// <param name="padding">Padding added around each obstacle while testing route intersections.</param>
        /// <param name="maxLanes">Number of positive and negative routing lanes to try on each axis.</param>
        public static VisioPage RouteConnectorsOrthogonalAroundShapes(this VisioPage page, double padding = 0.15D, int maxLanes = 12) {
            return page.RouteConnectorsOrthogonalAroundShapes(new VisioConnectorRoutingOptions {
                Padding = padding,
                MaxLanes = maxLanes
            });
        }

        /// <summary>
        /// Routes every connector on the page around unrelated top-level shapes using deterministic orthogonal lanes.
        /// </summary>
        /// <param name="page">Page whose connectors should be rerouted.</param>
        /// <param name="options">Routing options controlling padding, lane search, and whether zones/containers count as obstacles.</param>
        public static VisioPage RouteConnectorsOrthogonalAroundShapes(this VisioPage page, VisioConnectorRoutingOptions options) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            VisioConnectorRoutingOptions routingOptions = options;
            if (options.AvoidConnectorCrossings && options.ConnectorCrossingReferences == null) {
                routingOptions = options.Clone();
                routingOptions.ConnectorCrossingReferences = page.Connectors;
            }

            foreach (VisioConnector connector in page.Connectors) {
                connector.RouteOrthogonalAroundShapes(page.Shapes, routingOptions);
            }

            return page;
        }

        /// <summary>
        /// Removes explicit connector waypoints and returns the connector to dynamic routing.
        /// </summary>
        /// <param name="connector">Connector to reset.</param>
        public static VisioConnector ClearRoute(this VisioConnector connector) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            connector.Waypoints.Clear();
            connector.Kind = ConnectorKind.Dynamic;
            connector.PreservedGeometrySections.Clear();
            return connector;
        }

        /// <summary>
        /// Places connector text along the connector path.
        /// </summary>
        /// <param name="connector">Connector whose label should be placed.</param>
        /// <param name="position">Position along the connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public static VisioConnector PlaceLabel(this VisioConnector connector, double position = 0.5D, double offsetX = 0D, double offsetY = 0D, double width = 1.25D, double height = 0.3D) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            connector.LabelPlacement = VisioConnectorLabelPlacement.Along(position, offsetX, offsetY, width, height);
            return connector;
        }

        /// <summary>
        /// Places connector text at an absolute page coordinate.
        /// </summary>
        /// <param name="connector">Connector whose label should be placed.</param>
        /// <param name="pinX">Text pin X coordinate.</param>
        /// <param name="pinY">Text pin Y coordinate.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public static VisioConnector PlaceLabelAt(this VisioConnector connector, double pinX, double pinY, double width = 1.25D, double height = 0.3D) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            connector.LabelPlacement = VisioConnectorLabelPlacement.At(pinX, pinY, width, height);
            return connector;
        }

        /// <summary>
        /// Applies explicit waypoints to every selected connector.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="waypoints">Absolute page coordinates between start and end.</param>
        public static VisioConnectorSelection RouteThrough(this VisioConnectorSelection selection, params VisioConnectorWaypoint[] waypoints) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.RouteThrough(waypoints);
            }

            return selection;
        }

        /// <summary>
        /// Applies a generated orthogonal route to every selected connector.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="style">Orthogonal route orientation.</param>
        /// <param name="offset">Optional offset applied to the center routing lane.</param>
        public static VisioConnectorSelection RouteOrthogonal(this VisioConnectorSelection selection, VisioConnectorRouteStyle style = VisioConnectorRouteStyle.Auto, double offset = 0D) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.RouteOrthogonal(style, offset);
            }

            return selection;
        }

        /// <summary>
        /// Applies obstacle-aware orthogonal routing to every selected connector.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="obstacles">Shapes that selected connectors should avoid.</param>
        /// <param name="padding">Padding added around each obstacle while testing route intersections.</param>
        /// <param name="maxLanes">Number of positive and negative routing lanes to try on each axis.</param>
        public static VisioConnectorSelection RouteOrthogonalAroundShapes(this VisioConnectorSelection selection, IEnumerable<VisioShape> obstacles, double padding = 0.15D, int maxLanes = 12) {
            return selection.RouteOrthogonalAroundShapes(obstacles, new VisioConnectorRoutingOptions {
                Padding = padding,
                MaxLanes = maxLanes
            });
        }

        /// <summary>
        /// Applies obstacle-aware orthogonal routing to every selected connector.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="obstacles">Shapes that selected connectors should avoid.</param>
        /// <param name="options">Routing options controlling padding, lane search, and whether zones/containers count as obstacles.</param>
        public static VisioConnectorSelection RouteOrthogonalAroundShapes(this VisioConnectorSelection selection, IEnumerable<VisioShape> obstacles, VisioConnectorRoutingOptions options) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            if (options == null) {
                throw new ArgumentNullException(nameof(options));
            }

            VisioConnectorRoutingOptions routingOptions = options;
            if (options.AvoidConnectorCrossings && options.ConnectorCrossingReferences == null) {
                routingOptions = options.Clone();
                routingOptions.ConnectorCrossingReferences = selection;
            }

            foreach (VisioConnector connector in selection) {
                connector.RouteOrthogonalAroundShapes(obstacles, routingOptions);
            }

            return selection;
        }

        /// <summary>
        /// Removes explicit connector routes from every selected connector.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        public static VisioConnectorSelection ClearRoutes(this VisioConnectorSelection selection) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.ClearRoute();
            }

            return selection;
        }

        /// <summary>
        /// Places connector text along every selected connector path.
        /// </summary>
        /// <param name="selection">Connector selection.</param>
        /// <param name="position">Position along each connector path, from 0.0 to 1.0.</param>
        /// <param name="offsetX">Horizontal page-coordinate offset.</param>
        /// <param name="offsetY">Vertical page-coordinate offset.</param>
        /// <param name="width">Label text box width in page units.</param>
        /// <param name="height">Label text box height in page units.</param>
        public static VisioConnectorSelection PlaceLabels(this VisioConnectorSelection selection, double position = 0.5D, double offsetX = 0D, double offsetY = 0D, double width = 1.25D, double height = 0.3D) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            foreach (VisioConnector connector in selection) {
                connector.PlaceLabel(position, offsetX, offsetY, width, height);
            }

            return selection;
        }

        private static void ResolveEndpoint(VisioShape shape, VisioShape other, VisioConnectionPoint? connectionPoint, out double x, out double y) {
            if (connectionPoint != null) {
                (x, y) = shape.GetAbsolutePoint(connectionPoint.X, connectionPoint.Y);
                return;
            }

            (double left, double bottom, double right, double top) = shape.GetBounds();
            (double otherLeft, double otherBottom, double otherRight, double otherTop) = other.GetBounds();
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

        private static List<VisioShapeBounds> GetRoutingObstacleBounds(VisioConnector connector, IEnumerable<VisioShape> obstacles, double padding, VisioConnectorRoutingOptions options) {
            List<VisioShapeBounds> bounds = new();
            VisioShapeBounds fromBounds = connector.From.GetShapeBounds();
            VisioShapeBounds toBounds = connector.To.GetShapeBounds();
            foreach (VisioShape obstacle in obstacles) {
                if (ReferenceEquals(obstacle, connector.From) || ReferenceEquals(obstacle, connector.To)) {
                    continue;
                }

                if (obstacle.IsContainer && !options.IncludeContainers) {
                    continue;
                }

                if (obstacle.IsBackgroundSurface && !options.IncludeBackgroundSurfaces) {
                    continue;
                }

                if (obstacle.IsDiagramAdornment && !options.IncludeDiagramAdornments) {
                    continue;
                }

                VisioShapeBounds shapeBounds = obstacle.GetShapeBounds();
                if (!shapeBounds.IsEmpty &&
                    !Contains(shapeBounds, fromBounds) &&
                    !Contains(shapeBounds, toBounds)) {
                    bounds.Add(Inflate(shapeBounds, padding));
                }
            }

            return bounds;
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
            List<RoutePoint> points = new() {
                new RoutePoint(startX, startY)
            };

            foreach (VisioConnectorWaypoint waypoint in connector.Waypoints) {
                points.Add(new RoutePoint(waypoint.X, waypoint.Y));
            }

            points.Add(new RoutePoint(endX, endY));
            return points;
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
            return SegmentIntersectsBounds(route.Start, route.First, bounds) ||
                   SegmentIntersectsBounds(route.First, route.Second, bounds) ||
                   SegmentIntersectsBounds(route.Second, route.End, bounds);
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

        private static int CountConnectorCrossings(RouteCandidate candidate, IReadOnlyList<IReadOnlyList<RoutePoint>> connectorReferencePaths) {
            return CountConnectorCrossings(new[] { candidate.Start, candidate.First, candidate.Second, candidate.End }, connectorReferencePaths);
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
            return SegmentsIntersect(p1, p2, q1, q2) &&
                   !PointsEqual(p1, q1) &&
                   !PointsEqual(p1, q2) &&
                   !PointsEqual(p2, q1) &&
                   !PointsEqual(p2, q2);
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

        private static bool PointInside(RoutePoint point, VisioShapeBounds bounds) {
            return point.X > bounds.Left && point.X < bounds.Right &&
                   point.Y > bounds.Bottom && point.Y < bounds.Top;
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

        private static bool IsZero(double value) {
            return Math.Abs(value) < 1e-9;
        }

        private static VisioShapeBounds Inflate(VisioShapeBounds bounds, double padding) {
            return new VisioShapeBounds(
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

        private readonly struct RouteCandidate {
            public RouteCandidate(RoutePoint start, RoutePoint first, RoutePoint second, RoutePoint end)
                : this(start, first, second, end, new RouteScore(int.MaxValue, int.MaxValue, double.PositiveInfinity)) {
            }

            private RouteCandidate(RoutePoint start, RoutePoint first, RoutePoint second, RoutePoint end, RouteScore score) {
                Start = start;
                First = first;
                Second = second;
                End = end;
                Score = score;
            }

            public RoutePoint Start { get; }

            public RoutePoint First { get; }

            public RoutePoint Second { get; }

            public RoutePoint End { get; }

            public RouteScore Score { get; }

            public double Length => Distance(Start, First) + Distance(First, Second) + Distance(Second, End);

            public RouteCandidate WithScore(RouteScore score) {
                return new RouteCandidate(Start, First, Second, End, score);
            }
        }

        private readonly struct RouteScore {
            public RouteScore(int intersections, int connectorCrossings, double length) {
                Intersections = intersections;
                ConnectorCrossings = connectorCrossings;
                Length = length;
            }

            public int Intersections { get; }

            public int ConnectorCrossings { get; }

            public double Length { get; }

            public bool HasConflicts => Intersections > 0 || ConnectorCrossings > 0;

            public bool IsBetterThan(RouteScore other) {
                if (Intersections != other.Intersections) {
                    return Intersections < other.Intersections;
                }

                if (ConnectorCrossings != other.ConnectorCrossings) {
                    return ConnectorCrossings < other.ConnectorCrossings;
                }

                return Length < other.Length - 1e-9;
            }
        }

        private static double Distance(RoutePoint from, RoutePoint to) {
            double dx = to.X - from.X;
            double dy = to.Y - from.Y;
            return Math.Sqrt((dx * dx) + (dy * dy));
        }
    }
}
