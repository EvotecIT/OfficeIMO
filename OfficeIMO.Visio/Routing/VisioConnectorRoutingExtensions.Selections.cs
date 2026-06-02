using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    public static partial class VisioConnectorRoutingExtensions {

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
    }
}
