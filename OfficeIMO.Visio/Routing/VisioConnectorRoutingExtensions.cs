using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helpers for deterministic connector routing.
    /// </summary>
    public static partial class VisioConnectorRoutingExtensions {
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

            if (options.PageOptimizationPasses < 1) {
                throw new ArgumentOutOfRangeException(nameof(options), "Page optimization pass count must be at least one.");
            }

            ResolveEndpoint(connector.From, connector.To, connector.FromConnectionPoint, out double startX, out double startY);
            ResolveEndpoint(connector.To, connector.From, connector.ToConnectionPoint, out double endX, out double endY);
            IEnumerable<VisioShape> routingObstacles = options.IncludeGroupChildren
                ? ExpandRoutingObstacles(obstacles)
                : obstacles;
            List<VisioShapeBounds> obstacleBounds = GetRoutingObstacleBounds(connector, routingObstacles, padding, options);
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

            return connector.RouteThrough(resolved.Waypoints.Select(point => new VisioConnectorWaypoint(point.X, point.Y)));
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

            if (options.PageOptimizationPasses < 1) {
                throw new ArgumentOutOfRangeException(nameof(options), "Page optimization pass count must be at least one.");
            }

            List<VisioConnector> connectors = page.Connectors.ToList();
            VisioConnectorRoutingOptions routingOptions = options;
            if (options.AvoidConnectorCrossings && options.ConnectorCrossingReferences == null) {
                routingOptions = options.Clone();
                routingOptions.ConnectorCrossingReferences = connectors;
            }

            int passCount = routingOptions.AvoidConnectorCrossings
                ? routingOptions.PageOptimizationPasses
                : 1;
            for (int pass = 0; pass < passCount; pass++) {
                RouteScore before = ScorePageRoutes(connectors, page.Shapes, routingOptions);
                IReadOnlyList<VisioConnector> orderedConnectors = routingOptions.AvoidConnectorCrossings
                    ? OrderConnectorsForPageRouting(connectors, page.Shapes, routingOptions)
                    : connectors;
                foreach (VisioConnector connector in orderedConnectors) {
                    connector.RouteOrthogonalAroundShapes(page.Shapes, routingOptions);
                }

                if (!routingOptions.AvoidConnectorCrossings) {
                    break;
                }

                RouteScore after = ScorePageRoutes(connectors, page.Shapes, routingOptions);
                if (!after.HasConflicts || !after.IsBetterThan(before)) {
                    break;
                }
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
    }
}
