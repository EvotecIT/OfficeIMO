using System;

namespace OfficeIMO.Visio.Fluent {
    public partial class VisioFluentPage {
        /// <summary>
        /// Selects shapes with a matching shape name and configures them.
        /// </summary>
        public VisioFluentPage ShapesByName(string name, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectByName(name, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching universal shape name and configures them.
        /// </summary>
        public VisioFluentPage ShapesByNameU(string nameU, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectByNameU(nameU, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes created from a matching master universal name and configures them.
        /// </summary>
        public VisioFluentPage ShapesByMaster(string masterNameU, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectByMaster(masterNameU, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects Visio-native container shapes and configures them.
        /// </summary>
        public VisioFluentPage Containers(Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectContainers(), configure);
            return this;
        }

        /// <summary>
        /// Selects OfficeIMO callout or annotation shapes and configures them.
        /// </summary>
        public VisioFluentPage Callouts(Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectCallouts(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes that contain a Visio User cell and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithUserCell(string name, Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectWithUserCell(name), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching Visio User cell value and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithUserCell(string name, string value, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectWithUserCell(name, value, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with at least one hyperlink and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithHyperlinks(Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectWithHyperlinks(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with a matching hyperlink address and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithHyperlink(string address, Action<VisioShapeSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureShapes(Page.SelectWithHyperlink(address, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes with any explicit protection cell and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithProtection(Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectWithProtection(), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose protection state matches a predicate and configures them.
        /// </summary>
        public VisioFluentPage ShapesWithProtection(Func<VisioShapeProtection, bool> predicate, Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectWithProtection(predicate), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose bounds intersect the provided page-coordinate bounds and configures them.
        /// </summary>
        public VisioFluentPage ShapesIntersecting(VisioShapeBounds bounds, Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectIntersecting(bounds), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes whose bounds intersect a reference shape and configures them.
        /// </summary>
        public VisioFluentPage ShapesIntersecting(string shapeId, Action<VisioShapeSelection> configure, bool includeSelf = false) {
            ConfigureShapes(Page.SelectIntersecting(ResolveShape(shapeId), includeSelf), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes fully contained by the provided page-coordinate bounds and configures them.
        /// </summary>
        public VisioFluentPage ShapesContainedIn(VisioShapeBounds bounds, Action<VisioShapeSelection> configure) {
            ConfigureShapes(Page.SelectContainedIn(bounds), configure);
            return this;
        }

        /// <summary>
        /// Selects shapes fully contained by a reference shape and configures them.
        /// </summary>
        public VisioFluentPage ShapesContainedIn(string containerId, Action<VisioShapeSelection> configure, bool includeContainer = false) {
            ConfigureShapes(Page.SelectContainedIn(ResolveShape(containerId), includeContainer), configure);
            return this;
        }

        /// <summary>
        /// Selects every shape reachable from the provided shape through connectors and configures them.
        /// </summary>
        public VisioFluentPage ConnectedComponent(string shapeId, Action<VisioShapeSelection> configure, bool includeStart = true) {
            ConfigureShapes(Page.SelectConnectedComponent(ResolveShape(shapeId), includeStart), configure);
            return this;
        }

        /// <summary>
        /// Selects the shortest shape path between two connected shapes and configures it.
        /// </summary>
        public VisioFluentPage PathBetween(string fromId, string toId, Action<VisioShapeSelection> configure, bool includeEndpoints = true) {
            ConfigureShapes(Page.SelectPathBetween(ResolveShape(fromId), ResolveShape(toId), includeEndpoints), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors that start at the provided shape and configures them.
        /// </summary>
        public VisioFluentPage OutgoingConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectOutgoingConnectors(ResolveShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors that end at the provided shape and configures them.
        /// </summary>
        public VisioFluentPage IncomingConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectIncomingConnectors(ResolveShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors attached to the provided shape and configures them.
        /// </summary>
        public VisioFluentPage ConnectedConnectors(string shapeId, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectConnectedConnectors(ResolveShape(shapeId)), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors assigned to a page layer and configures them.
        /// </summary>
        public VisioFluentPage ConnectorsInLayer(string layerName, Action<VisioConnectorSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureConnectors(Page.SelectConnectorsInLayer(layerName, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with at least one hyperlink and configures them.
        /// </summary>
        public VisioFluentPage ConnectorsWithHyperlinks(Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectConnectorsWithHyperlinks(), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with a matching hyperlink address and configures them.
        /// </summary>
        public VisioFluentPage ConnectorsWithHyperlink(string address, Action<VisioConnectorSelection> configure, StringComparison comparison = StringComparison.OrdinalIgnoreCase) {
            ConfigureConnectors(Page.SelectConnectorsWithHyperlink(address, comparison), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors with any explicit protection cell and configures them.
        /// </summary>
        public VisioFluentPage ConnectorsWithProtection(Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectConnectorsWithProtection(), configure);
            return this;
        }

        /// <summary>
        /// Selects connectors whose protection state matches a predicate and configures them.
        /// </summary>
        public VisioFluentPage ConnectorsWithProtection(Func<VisioProtection, bool> predicate, Action<VisioConnectorSelection> configure) {
            ConfigureConnectors(Page.SelectConnectorsWithProtection(predicate), configure);
            return this;
        }

        private VisioShape ResolveShape(string shapeId) {
            if (string.IsNullOrWhiteSpace(shapeId)) {
                throw new ArgumentException("Shape id cannot be null or whitespace.", nameof(shapeId));
            }

            if (!_byId.TryGetValue(shapeId, out VisioShape? shape)) {
                throw new ArgumentException($"Unknown shape id '{shapeId}'.", nameof(shapeId));
            }

            return shape;
        }
    }
}
