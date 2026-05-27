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
    }
}
